/*
 * Meeting Hours Guard
 * Warns (and on Send, prompts) organizers when a meeting falls outside
 * company meeting hours. Edit POLICY to change the rule.
 */

const POLICY = {
  timeZone: "Europe/Istanbul",           // IANA name used for the check
  days: ["Mon", "Tue", "Wed", "Thu", "Fri"],
  start: 1000,                           // HHmm, earliest allowed start
  end: 1830,                             // HHmm, latest allowed end
  label: "Mon to Fri, 10:00 to 18:30 Istanbul time",
};

const NOTICE_KEY = "meetingHoursGuard";

function localParts(date) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: POLICY.timeZone,
    weekday: "short",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date);
  const p = {};
  for (const x of parts) p[x.type] = x.value;
  return {
    day: p.weekday,
    date: `${p.year}-${p.month}-${p.day}`,
    hhmm: Number(p.hour + p.minute),
  };
}

function fmt(hhmm) {
  const s = String(hhmm).padStart(4, "0");
  return `${s.slice(0, 2)}:${s.slice(2)}`;
}

/* Returns null when the meeting is inside hours, else a short reason. */
function checkTimes(start, end) {
  const s = localParts(start);
  const e = localParts(end);
  if (!POLICY.days.includes(s.day)) return "falls on a weekend";
  if (s.hhmm < POLICY.start) return `starts before ${fmt(POLICY.start)}`;
  if (s.date !== e.date || e.hhmm > POLICY.end) return `ends after ${fmt(POLICY.end)}`;
  return null;
}

/* ---------- Office glue (skipped when run under Node for tests) ---------- */

function log(...args) {
  console.log("[MHG]", ...args);
}
log("commands.js loaded", new Date().toISOString());

/* getAsync as a promise, with a timeout so a silent host never hangs the handler. */
function getAsyncP(prop, name, ms = 4000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${name}.getAsync timed out`)), ms);
    try {
      prop.getAsync((r) => {
        clearTimeout(timer);
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(r.error || new Error(`${name}.getAsync failed`));
      });
    } catch (e) {
      clearTimeout(timer);
      reject(e);
    }
  });
}

function callAsyncP(fn) {
  return new Promise((resolve) => fn((r) => resolve(r)));
}

/* Reads one property at a time (sequential, so a stuck call is easy to spot in the log). */
async function evaluate(item) {
  const start = await getAsyncP(item.start, "start");
  const end = await getAsyncP(item.end, "end");
  log("times", start && start.toISOString(), end && end.toISOString());
  let allDay = false;
  if (item.isAllDayEvent) allDay = await getAsyncP(item.isAllDayEvent, "isAllDayEvent").catch(() => false);
  const required = await getAsyncP(item.requiredAttendees, "requiredAttendees").catch(() => []);
  const optional = await getAsyncP(item.optionalAttendees, "optionalAttendees").catch(() => []);
  log("allDay", allDay, "attendees", required.length + optional.length);
  if (allDay) return null;                                  // all-day events are out of scope
  if (required.length + optional.length === 0) return null; // personal appointment, no invite
  const reason = checkTimes(start, end);
  log("reason", reason);
  return reason;
}

function showNotice(item, message) {
  return callAsyncP((cb) =>
    item.notificationMessages.replaceAsync(
      NOTICE_KEY,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message.slice(0, 150),
        icon: "Icon.16x16",
        persistent: false,
      },
      cb
    )
  );
}

function clearNotice(item) {
  return callAsyncP((cb) => item.notificationMessages.removeAsync(NOTICE_KEY, cb));
}

/* OnNewAppointmentOrganizer, OnAppointmentTimeChanged, OnAppointmentAttendeesChanged */
async function onAppointmentChangedHandler(event) {
  log("changed handler", event && event.type);
  const item = Office.context.mailbox.item;
  try {
    const reason = await evaluate(item);
    if (reason) {
      await showNotice(item, `Outside meeting hours: this meeting ${reason}. Meeting hours are ${POLICY.label}.`);
    } else {
      await clearNotice(item);
    }
  } catch (e) {
    console.error("Meeting Hours Guard:", e);
  }
  event.completed();
}

/* OnAppointmentSend (Smart Alerts) */
async function onAppointmentSendHandler(event) {
  log("send handler start");
  try {
    const reason = await evaluate(Office.context.mailbox.item);
    if (!reason) {
      event.completed({ allowEvent: true });
      return;
    }
    event.completed({
      allowEvent: false,
      errorMessage:
        `This meeting ${reason}. Company meeting hours are ${POLICY.label}. ` +
        `Please move it into that window, or choose Send Anyway if it can't move.`,
    });
  } catch (e) {
    console.error("Meeting Hours Guard:", e);
    event.completed({ allowEvent: true }); // fail open: never trap a send on an add-in error
  }
}

/* Ribbon button "Check hours" */
async function checkMeetingHours(event) {
  log("button handler");
  const item = Office.context.mailbox.item;
  try {
    const reason = await evaluate(item);
    await showNotice(
      item,
      reason
        ? `Outside meeting hours: this meeting ${reason}. Meeting hours are ${POLICY.label}.`
        : `Within meeting hours (${POLICY.label}).`
    );
  } catch (e) {
    console.error("Meeting Hours Guard:", e);
  }
  event.completed();
}

/* Map manifest FunctionName values to handlers. Register at load time (as Microsoft's samples do)
   and again once Office is ready, in case Office.actions isn't available yet when this script runs. */
let associated = false;
function register(when) {
  if (associated) return;
  if (!(Office.actions && Office.actions.associate)) {
    log("Office.actions not available", when);
    return;
  }
  Office.actions.associate("onAppointmentChangedHandler", onAppointmentChangedHandler);
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
  Office.actions.associate("checkMeetingHours", checkMeetingHours);
  associated = true;
  log("handlers associated", when);
}

if (typeof Office !== "undefined") {
  try {
    register("at load");
  } catch (e) {
    log("associate at load failed", e && e.message);
  }
  Office.onReady((info) => {
    log("Office ready", info && info.host, info && info.platform);
    register("on ready");
  });
}

if (typeof module !== "undefined") {
  module.exports = { POLICY, checkTimes, localParts };
}
