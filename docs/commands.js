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

function getAsyncP(prop) {
  return new Promise((resolve, reject) => {
    prop.getAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(r.error);
    });
  });
}

function callAsyncP(fn) {
  return new Promise((resolve) => fn((r) => resolve(r)));
}

async function evaluate(item) {
  const [start, end, allDay, required, optional] = await Promise.all([
    getAsyncP(item.start),
    getAsyncP(item.end),
    item.isAllDayEvent ? getAsyncP(item.isAllDayEvent).catch(() => false) : false,
    getAsyncP(item.requiredAttendees).catch(() => []),
    getAsyncP(item.optionalAttendees).catch(() => []),
  ]);
  if (allDay) return null;                                  // all-day events are out of scope
  if (required.length + optional.length === 0) return null; // personal appointment, no invite
  return checkTimes(start, end);
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

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onAppointmentChangedHandler", onAppointmentChangedHandler);
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
  Office.actions.associate("checkMeetingHours", checkMeetingHours);
}

if (typeof module !== "undefined") {
  module.exports = { POLICY, checkTimes, localParts };
}
