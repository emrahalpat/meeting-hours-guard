// Unit tests for the meeting-hours rule (run: npm test). Times are given in UTC; Istanbul is UTC+3.
const assert = require("assert");
const { checkTimes } = require("./src/commands.js");

const cases = [
  // [label, startUTC, endUTC, expected]
  ["Wed 10:00-11:00 inside", "2026-11-11T07:00:00Z", "2026-11-11T08:00:00Z", null],
  ["Wed 17:30-18:30 ends exactly at limit", "2026-11-11T14:30:00Z", "2026-11-11T15:30:00Z", null],
  ["Wed 09:30-10:30 starts early", "2026-11-11T06:30:00Z", "2026-11-11T07:30:00Z", "starts before 10:00"],
  ["Wed 18:00-19:00 ends late", "2026-11-11T15:00:00Z", "2026-11-11T16:00:00Z", "ends after 18:30"],
  ["Wed 20:00-21:00 evening", "2026-11-11T17:00:00Z", "2026-11-11T18:00:00Z", "ends after 18:30"],
  ["Fri 18:00 to Mon 10:00 spans weekend", "2026-11-13T15:00:00Z", "2026-11-16T07:00:00Z", "ends after 18:30"],
  ["Sat 11:00-12:00 weekend", "2026-11-14T08:00:00Z", "2026-11-14T09:00:00Z", "falls on a weekend"],
  ["Sun 14:00-15:00 weekend", "2026-11-15T11:00:00Z", "2026-11-15T12:00:00Z", "falls on a weekend"],
  ["Mon 10:00-18:30 full day inside", "2026-11-16T07:00:00Z", "2026-11-16T15:30:00Z", null],
  // Public holidays and half days (2027)
  ["Fri 1 Jan 11:00 New Year's Day", "2027-01-01T08:00:00Z", "2027-01-01T09:00:00Z", "falls on a public holiday (New Year's Day)"],
  ["Wed 10 Mar 14:00 Ramadan Feast", "2027-03-10T11:00:00Z", "2027-03-10T12:00:00Z", "falls on a public holiday (Ramadan Feast)"],
  ["Wed 19 May Youth Day + Feast", "2027-05-19T08:00:00Z", "2027-05-19T09:00:00Z", "falls on a public holiday (Youth and Sports Day, Sacrifice Feast)"],
  ["Fri 29 Oct Republic Day", "2027-10-29T08:00:00Z", "2027-10-29T09:00:00Z", "falls on a public holiday (Republic Day)"],
  ["Thu 28 Oct 11:00-12:00 half day morning", "2027-10-28T08:00:00Z", "2027-10-28T09:00:00Z", null],
  ["Thu 28 Oct 12:00-13:00 ends at half-day limit", "2027-10-28T09:00:00Z", "2027-10-28T10:00:00Z", null],
  ["Thu 28 Oct 12:30-13:30 half day afternoon", "2027-10-28T09:30:00Z", "2027-10-28T10:30:00Z", "ends after 13:00 on a half day (Republic Day Eve)"],
  ["Tue 9 Mar 15:00 Ramadan Feast Eve", "2027-03-09T12:00:00Z", "2027-03-09T13:00:00Z", "ends after 13:00 on a half day (Ramadan Feast Eve)"],
  ["Tue 9 Mar 09:00 Ramadan Feast Eve early", "2027-03-09T06:00:00Z", "2027-03-09T07:00:00Z", "starts before 10:00"],
  ["Wed 10 Nov Ataturk Commemoration is a work day", "2027-11-10T08:00:00Z", "2027-11-10T09:00:00Z", null],
  ["Fri 31 Dec New Year's Eve is a work day", "2027-12-31T08:00:00Z", "2027-12-31T09:00:00Z", null],
  ["Mon 8 Feb Ramadan start is a work day", "2027-02-08T08:00:00Z", "2027-02-08T09:00:00Z", null],
  ["Sat 15 May half day on a weekend", "2027-05-15T08:00:00Z", "2027-05-15T09:00:00Z", "falls on a weekend"],
];

let failed = 0;
for (const [label, s, e, want] of cases) {
  const got = checkTimes(new Date(s), new Date(e));
  try {
    assert.strictEqual(got, want);
    console.log(`ok    ${label}`);
  } catch {
    failed++;
    console.log(`FAIL  ${label}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
  }
}
// Every holiday notice must fit Outlook's 150-character info bar.
const { POLICY, holidayYearCovered, noticeText } = require("./src/commands.js");
const reasons = [
  ...Object.values(POLICY.holidays).map((n) => `falls on a public holiday (${n})`),
  ...Object.values(POLICY.halfDays).map((n) => `ends after 13:00 on a half day (${n})`),
];
for (const r of reasons) {
  const msg = noticeText(r);
  if (msg.length > 150) { failed++; console.log(`FAIL  notice too long (${msg.length}): ${msg}`); }
}
const longest = Math.max(...reasons.map((r) => noticeText(r).length));
console.log(`ok    longest holiday notice is ${longest} of 150 characters`);
if (!holidayYearCovered(new Date("2027-06-01T09:00:00Z"))) { failed++; console.log("FAIL  2027 holiday list missing"); }

console.log(failed ? `\n${failed} failed` : `\nall ${cases.length} cases passed`);
process.exit(failed ? 1 : 0);
