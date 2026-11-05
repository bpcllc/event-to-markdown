/*
 * Event to Markdown - Outlook add-in task pane
 * Reads the open calendar item (read or compose mode) and renders it as
 * Markdown or plain text for copying. Read-only: never changes the item.
 */
"use strict";

/* ---------- Data collection (Office.js) ---------- */

// Read-mode fields are plain values; compose-mode fields expose getAsync().
function getAsyncValue(obj) {
  return new Promise(function (resolve) {
    try {
      obj.getAsync(function (r) {
        resolve(r && r.status === Office.AsyncResultStatus.Succeeded ? r.value : null);
      });
    } catch (e) {
      resolve(null);
    }
  });
}

async function readField(item, name) {
  var v;
  try { v = item[name]; } catch (e) { return null; }
  if (v === undefined || v === null) return null;
  if (typeof v === "object" && typeof v.getAsync === "function" && !(v instanceof Date)) {
    return getAsyncValue(v);
  }
  return v;
}

function readBody(item) {
  return new Promise(function (resolve) {
    try {
      item.body.getAsync(Office.CoercionType.Text, function (r) {
        resolve(r && r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
      });
    } catch (e) {
      resolve("");
    }
  });
}

// Convert a UTC Date to Outlook's configured client time zone.
// Offset is derived from the returned wall-clock fields, so no sign ambiguity.
function toClientLocal(date) {
  if (!(date instanceof Date) || isNaN(date)) return null;
  try {
    var lct = Office.context.mailbox.convertToLocalClientTime(date);
    var wallAsUtc = Date.UTC(lct.year, lct.month, lct.date, lct.hours, lct.minutes, lct.seconds || 0);
    var offsetMin = Math.round((wallAsUtc - Math.floor(date.getTime() / 1000) * 1000) / 60000);
    return { y: lct.year, mo: lct.month + 1, d: lct.date, h: lct.hours, mi: lct.minutes, offsetMin: offsetMin };
  } catch (e) {
    return {
      y: date.getFullYear(), mo: date.getMonth() + 1, d: date.getDate(),
      h: date.getHours(), mi: date.getMinutes(), offsetMin: -date.getTimezoneOffset()
    };
  }
}

async function collectEvent() {
  var item = Office.context.mailbox.item;
  if (!item) throw new Error("No item is open.");

  var results = await Promise.all([
    readField(item, "subject"),
    readField(item, "start"),
    readField(item, "end"),
    readField(item, "location"),
    readField(item, "organizer"),
    readField(item, "requiredAttendees"),
    readField(item, "optionalAttendees"),
    readField(item, "recurrence"),
    readBody(item)
  ]);

  var start = results[1] instanceof Date ? results[1] : (results[1] ? new Date(results[1]) : null);
  var end = results[2] instanceof Date ? results[2] : (results[2] ? new Date(results[2]) : null);

  return {
    subject: results[0] || "",
    start: start,
    end: end,
    startLocal: toClientLocal(start),
    endLocal: toClientLocal(end),
    location: results[3] || "",
    organizer: results[4] || null,
    required: Array.isArray(results[5]) ? results[5] : [],
    optional: Array.isArray(results[6]) ? results[6] : [],
    recurrenceType: results[7] && results[7].recurrenceType ? String(results[7].recurrenceType) : "",
    body: results[8] || ""
  };
}

/* ---------- Formatting (pure functions, unit-testable) ---------- */

function pad(n) { return (n < 10 ? "0" : "") + n; }

function fmtLocalDate(l) { return l.y + "-" + pad(l.mo) + "-" + pad(l.d); }
function fmtLocalTime(l) { return pad(l.h) + ":" + pad(l.mi); }

function fmtOffset(min) {
  var sign = min < 0 ? "-" : "+";
  var a = Math.abs(min);
  return "UTC" + sign + pad(Math.floor(a / 60)) + ":" + pad(a % 60);
}

function fmtWhen(ev) {
  var s = ev.startLocal, e = ev.endLocal;
  if (!s) return "";
  if (!e) return fmtLocalDate(s) + " " + fmtLocalTime(s) + " (" + fmtOffset(s.offsetMin) + ")";
  var sameDay = s.y === e.y && s.mo === e.mo && s.d === e.d;
  var tz = s.offsetMin === e.offsetMin
    ? fmtOffset(s.offsetMin)
    : fmtOffset(s.offsetMin) + " / " + fmtOffset(e.offsetMin);
  return sameDay
    ? fmtLocalDate(s) + " " + fmtLocalTime(s) + "–" + fmtLocalTime(e) + " (" + tz + ")"
    : fmtLocalDate(s) + " " + fmtLocalTime(s) + " – " + fmtLocalDate(e) + " " + fmtLocalTime(e) + " (" + tz + ")";
}

function isoUtc(d) {
  return d instanceof Date && !isNaN(d) ? d.toISOString().replace(/\.\d{3}Z$/, "Z") : "";
}

function fmtUtc(ev) {
  var a = isoUtc(ev.start), b = isoUtc(ev.end);
  return a && b ? a + " – " + b : a;
}

function fmtPerson(p) {
  if (!p) return "";
  var name = p.displayName || "";
  var email = p.emailAddress || "";
  if (name && email && name.toLowerCase() !== email.toLowerCase()) return name + " (" + email + ")";
  return name || email;
}

var RESPONSE_LABELS = {
  accepted: "Accepted", declined: "Declined", tentative: "Tentative",
  organizer: "Organizer", none: "No response"
};

function fmtResponse(p) {
  var r = p && p.appointmentResponse ? String(p.appointmentResponse).toLowerCase() : "";
  return RESPONSE_LABELS[r] || "";
}

// Escape characters that would change meaning in inline Markdown.
function mdEscape(s) {
  return String(s).replace(/([\\`*_\[\]<>])/g, "\\$1");
}

function cleanBody(text) {
  return String(text || "")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function toMarkdown(ev, opt) {
  var L = [];
  L.push("## " + (ev.subject ? mdEscape(ev.subject) : "(no subject)"));
  L.push("");
  var when = fmtWhen(ev);
  if (when) L.push("- **When:** " + when);
  if (opt.utc && ev.start) L.push("- **UTC:** " + fmtUtc(ev));
  if (ev.location) L.push("- **Location:** " + mdEscape(ev.location));
  if (ev.organizer) L.push("- **Organizer:** " + mdEscape(fmtPerson(ev.organizer)));
  if (ev.recurrenceType) L.push("- **Recurrence:** " + ev.recurrenceType);

  if (opt.attendees) {
    [["Required attendees", ev.required], ["Optional attendees", ev.optional]].forEach(function (g) {
      if (!g[1].length) return;
      L.push("- **" + g[0] + ":**");
      g[1].forEach(function (p) {
        var resp = opt.responses ? fmtResponse(p) : "";
        L.push("  - " + mdEscape(fmtPerson(p)) + (resp ? " — " + resp : ""));
      });
    });
  }

  var body = opt.body ? cleanBody(ev.body) : "";
  if (body) {
    L.push("");
    L.push("### Notes");
    L.push("");
    L.push(body);
  }
  return L.join("\n") + "\n";
}

function toPlainText(ev, opt) {
  var L = [];
  function field(label, value) { if (value) L.push((label + ":").padEnd(12) + value); }
  field("Subject", ev.subject || "(no subject)");
  field("When", fmtWhen(ev));
  if (opt.utc) field("UTC", fmtUtc(ev));
  field("Location", ev.location);
  field("Organizer", fmtPerson(ev.organizer));
  field("Recurrence", ev.recurrenceType);

  if (opt.attendees) {
    [["Required attendees", ev.required], ["Optional attendees", ev.optional]].forEach(function (g) {
      if (!g[1].length) return;
      L.push(g[0] + ":");
      g[1].forEach(function (p) {
        var resp = opt.responses ? fmtResponse(p) : "";
        L.push("  " + fmtPerson(p) + (resp ? " - " + resp : ""));
      });
    });
  }

  var body = opt.body ? cleanBody(ev.body) : "";
  if (body) {
    L.push("");
    L.push("Notes:");
    L.push(body);
  }
  return L.join("\n") + "\n";
}

/* ---------- UI ---------- */

var currentEvent = null;
var PREF_KEY = "e2m.prefs.v1";

function $(id) { return document.getElementById(id); }

function getOptions() {
  var fmt = document.querySelector('input[name="format"]:checked');
  return {
    format: fmt ? fmt.value : "markdown",
    attendees: $("optAttendees").checked,
    responses: $("optResponses").checked,
    utc: $("optUtc").checked,
    body: $("optBody").checked
  };
}

function savePrefs() {
  try { localStorage.setItem(PREF_KEY, JSON.stringify(getOptions())); } catch (e) { /* storage unavailable */ }
}

function loadPrefs() {
  var p = null;
  try { p = JSON.parse(localStorage.getItem(PREF_KEY) || "null"); } catch (e) { p = null; }
  if (!p) return;
  var radio = document.querySelector('input[name="format"][value="' + (p.format === "text" ? "text" : "markdown") + '"]');
  if (radio) radio.checked = true;
  ["Attendees", "Responses", "Utc", "Body"].forEach(function (k) {
    var key = k.charAt(0).toLowerCase() + k.slice(1);
    if (typeof p[key] === "boolean") $("opt" + k).checked = p[key];
  });
}

function setStatus(msg, kind) {
  var el = $("status");
  el.textContent = msg || "";
  el.className = kind || "";
}

function render() {
  savePrefs();
  $("optResponses").disabled = !$("optAttendees").checked;
  if (!currentEvent) return;
  var opt = getOptions();
  $("output").value = opt.format === "text" ? toPlainText(currentEvent, opt) : toMarkdown(currentEvent, opt);
  setStatus("");
}

async function load() {
  setStatus("Reading event…");
  try {
    currentEvent = await collectEvent();
    render();
  } catch (e) {
    currentEvent = null;
    $("output").value = "";
    setStatus("Couldn't read this item: " + (e && e.message ? e.message : e), "warn");
  }
}

async function copyOutput() {
  var ta = $("output");
  var text = ta.value;
  var ok = false;

  // 1) execCommand runs synchronously inside the click gesture; works in the
  //    add-in iframe where the async Clipboard API is commonly blocked.
  ta.focus();
  ta.select();
  ta.setSelectionRange(0, text.length);
  try { ok = document.execCommand("copy"); } catch (e) { ok = false; }

  // 2) Fall back to the async Clipboard API if a host ever permits it.
  if (!ok && navigator.clipboard && navigator.clipboard.writeText) {
    try { await navigator.clipboard.writeText(text); ok = true; } catch (e) { ok = false; }
  }

  if (ok) {
    setStatus("Copied.", "ok");
  } else {
    ta.focus();
    ta.select();
    setStatus("Copy blocked — text is selected, press Ctrl+C.", "warn");
  }
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      setStatus("This add-in runs in Outlook.", "warn");
      return;
    }
    loadPrefs();
    document.querySelectorAll('input[name="format"], input[type="checkbox"]').forEach(function (el) {
      el.addEventListener("change", render);
    });
    $("copy").addEventListener("click", copyOutput);
    $("refresh").addEventListener("click", load);
    try {
      if (Office.context.mailbox.addHandlerAsync && Office.EventType.ItemChanged) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, load);
      }
    } catch (e) { /* not supported on this client */ }
    load();
  });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { toMarkdown: toMarkdown, toPlainText: toPlainText, fmtWhen: fmtWhen, fmtOffset: fmtOffset, mdEscape: mdEscape, cleanBody: cleanBody };
}
