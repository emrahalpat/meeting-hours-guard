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
console.log(failed ? `\n${failed} failed` : `\nall ${cases.length} passed`);
process.exit(failed ? 1 : 0);
