# Meeting Hours Guard (Outlook on the web, dev build)

Outlook Smart Alerts add-in. When someone organizes a meeting outside **Mon to Fri, 10:00 to 18:30 Istanbul time**:

- while drafting: an info bar appears ("Outside meeting hours: ...") as soon as the time or attendees put it out of hours
- on Send: a Smart Alerts dialog explains why, with **Send Anyway** (SendMode `PromptUser`)
- ribbon button **Check hours** runs the same check on demand

Personal appointments (no attendees) and all-day events are ignored. If the add-in errors, the meeting sends (fail open).

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Add-in only (XML) manifest; works on web, new/classic Windows and Mac |
| `src/commands.js` | Rule (`POLICY`) and the event handlers |
| `src/commands.html` | Runtime page Outlook on the web loads (office.js + commands.js) |
| `server.js` | HTTPS static server on https://localhost:3000 |
| `test.js` | Unit tests for the hours rule |

## Run it locally

```bash
npm run certs      # once: creates and trusts a localhost certificate (asks for your Mac password)
npm start          # serves https://localhost:3000, keep it running while testing
npm test           # rule tests
npm run validate   # manifest check (Microsoft's validator)
```

Check https://localhost:3000/commands.html opens in your browser without a certificate warning.

## Sideload in Outlook on the web

1. Open https://aka.ms/olksideload while signed in to Outlook on the web (opens **Add-ins for Outlook**).
2. **My add-ins** > **Custom Addins** > **Add a custom add-in** > **Add from File** > pick `manifest.xml` > Install.
3. Reload Outlook. It can take a few minutes before event-based add-ins start firing.

If the **Add a custom add-in** option is missing, the tenant's role assignment policy blocks custom add-ins (`My Custom Apps` role). Test with an account that has it, or deploy centrally (below).

## Test cases

Calendar > **New event** (full form), add at least one attendee:

| Set the time to | Expect |
|---|---|
| Wed 11:00 to 12:00 | no bar, sends normally |
| Wed 09:00 to 09:30 | bar "starts before 10:00"; Send shows the dialog |
| Wed 18:00 to 19:00 | bar "ends after 18:30"; Send shows the dialog |
| Saturday any time | bar "falls on a weekend"; Send shows the dialog |

Also try the quick-create card (click a time slot in the calendar grid) and forwarding a meeting. Microsoft documents that `OnAppointmentSend` only fires when the meeting is sent from a separate window, so those paths may skip the check.

Debug: browser DevTools console (the handlers log `Meeting Hours Guard:` on errors).

## Change the rule

Edit `POLICY` at the top of `src/commands.js` (time zone, days, start/end as HHmm, label), then `npm test`. To stop people overriding: change `SendMode="PromptUser"` in `manifest.xml` to `SoftBlock` (must fix, sends if the add-in is down) or `Block` (must fix, nothing sends if the add-in is down; admin deployment only).

## Production deployment

1. Host `src/` and `assets/` on any HTTPS static host (e.g. Azure Static Web Apps).
2. Replace every `https://localhost:3000` in `manifest.xml` with that host; bump `<Version>`.
3. Microsoft 365 admin center > Settings > Integrated apps > Upload custom apps > Office Add-in > upload `manifest.xml`, assign users or groups. Event-based add-ins only run when deployed by an admin in production.
