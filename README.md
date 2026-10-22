# ECS Staff Portal

Google Apps Script web app for EC Smart Automations staff: personal To Do with due dates and in-app reminders, Daily Reflection, Notice Board (announcements and circulars), and a Document Library with per-person or per-role visibility.

- **Database:** Google Sheet `ECS Staff Portal - DB` (`10GJ5INILBMX14bPidnkLHyndGRurdDDYUq6tcDbrhcs`)
- **Files:** `Code.js` (server), `Index.html` (page), `appsscript.json` (manifest, timezone Asia/Kolkata)

## One-time setup

1. In the Drive folder for this codebase:
   ```
   git clone https://github.com/hariofficial94/ecs-staff-portal .
   clasp create --type standalone --title "ECS Staff Portal"
   ```
   `clasp create` writes `.clasp.json` (ignored by git). If the Apps Script project already exists, use `clasp clone <scriptId>` into an empty folder instead, then copy `.clasp.json` here.
2. `clasp push`
3. In the Apps Script editor, run `setupCheck` once. Approve the Sheets permission. The log should say "All tabs and headers OK".
4. **Deploy > New deployment > Web app.** Execute as: *Me*. Who has access: *Anyone*. Copy the URL and share it with staff.

## Every change after that

```
git pull
clasp push
```
Then **Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.** Saving or pushing alone does not change the live app. Do not create a second deployment unless you want a second URL.

## Feature flags

At the top of `Code.js`:

```javascript
const FEATURES_ENABLED = {
  todo: true, reflection: true, announcements: true, library: true,
  assignedTasks: false, teamCompletion: false
};
```

Flipping a flag is the whole switch: the page renders a pane only when its flag is `true`, and every server function for that feature refuses to run when it is `false`. Push and deploy a new version after changing it.

## Admin

Whoever has an empty `Reports_To` in `Config_EC` gets the Admin tab (posting announcements, adding documents, read receipts). This is derived from the hierarchy, so a handover only needs a Config change.

## Adding a staff member

1. In the Apps Script editor, run `generateStaffCode` and copy the code from the log.
2. Add a row in `Config_EC`: Role (unique), Name, PIN, Type, `Active`, the new code, Reports_To (a Role, not a name).

Set `Staff_Status` to `Placeholder` to keep someone on the roster without letting them sign in or appear in reports.

## Documents

Share each Drive file as "Anyone with the link can view". The portal only sends a document's link to the people in its audience (and to their managers up the reporting line), but anyone who has the link can forward it.

## Design notes

- PINs are checked on the server only; the browser holds a random session token (6 hours, refreshed on use). Five wrong PINs lock that name for 15 minutes.
- All Sheet columns are plain text (except `Hours`), and `readTable_` formats any date cell to a string, so no `Date` object can reach `google.script.run`.
- Every row has a prefixed ID: `TD-<StaffCode>-0001`, `RP-<StaffCode>-YYYYMMDD`, `TL-00001`, `AN-0001`, `RD-000001`, `LIB-0001`, `AT-00001`, `TC-00001`.
- To-dos are one row for their whole life. Open items stay on the list until done or cancelled; nothing is copied day to day.
- Daily Reflection has no deadline, lock or lateness label. Any past date can be written or edited.
