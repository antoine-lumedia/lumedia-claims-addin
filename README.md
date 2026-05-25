# Claims Migration — task-pane add-in

A small Excel add-in that gives the proper dialog UX for manual status changes
on `ClaimsTracker.xlsx`. Adds a **Claims** ribbon tab with a single button:
**Apply Status Changes**.

## What it does

1. Reads pending status changes on the active sheet (rows where the user has
   set `status_action` to a value).
2. Shows a confirmation list. For any migration TO a `*_Stuck` sheet, a
   required `block_reason` textarea is shown — Apply stays disabled until all
   such reasons are at least 3 characters.
3. Posts one HTTP request to Flow 3 per migration.
4. Clears the dropdowns of successfully migrated contracts.
5. Shows a summary.

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Office Add-in manifest (registers the ribbon tab) |
| `taskpane.html` | UI markup |
| `taskpane.js` | Logic |
| `taskpane.css` | Styling |
| `icon-16.png`, `icon-32.png`, `icon-80.png` | Ribbon icons (provide your own; placeholder 1×1 PNGs included if needed) |

## Configuration

The add-in reads three URLs from named cells on the hidden `_FlowConfig` sheet
of `ClaimsTracker.xlsx`:

| Named cell | What to paste |
|---|---|
| `flow3_url` | HTTP trigger URL of Power Automate Flow 3 |
| `script_read_url` | Office Scripts REST URL for `readPendingMigrations` |
| `script_clear_url` | Office Scripts REST URL for `clearStatusActions` |

To get an Office Scripts REST URL:
1. Open the script in Excel Online.
2. Click the `...` menu in the script editor → **More options** → **API**.
3. Copy the run URL (looks like `https://officescripts.microsoftusercontent.com/.../run`).

To get the Flow 3 HTTP trigger URL:
1. Open Flow 3 in Power Automate.
2. Click on the `When a HTTP request is received` trigger.
3. Copy the `HTTP POST URL` (only available after the flow has been saved at
   least once).

Paste each into column B of the corresponding row on `_FlowConfig`. The
add-in re-reads them on every invocation, so you can change them any time.

## Install — tenant-wide (recommended)

1. Upload the four files (manifest, html, js, css) plus icons to
   `/Documents partages/Customer Claims/_AddIn/` on the finance SharePoint
   site.
2. Open **Microsoft 365 Admin Center** → **Settings** → **Integrated apps** →
   **Upload custom apps**.
3. Choose **Office Add-in**, upload `manifest.xml`.
4. Assign to users (the claims team) or to a security group.
5. After deployment finishes (up to 24h, usually minutes), users open
   `ClaimsTracker.xlsx` in Excel Online and see the **Claims** tab.

## Install — sideload for testing (single user)

1. Excel Online → Insert → Office Add-ins → Manage My Add-ins → Upload My
   Add-in.
2. Choose the `manifest.xml`.
3. Confirms only for the current user, scoped to the current workbook session.

## Icons

The manifest references three PNG icons:
- 16×16: for the small ribbon button
- 32×32: for the medium ribbon button
- 80×80: for the high-resolution version

You can use any logo. If you want a quick neutral placeholder, generate one
with the Lumedia colours or use any plain 32×32 PNG. The add-in works without
the icons — Office shows a generic placeholder.

## Notes on Office Scripts REST authentication

The add-in calls Office Scripts REST endpoints from the browser, using a
single sign-on token via `OfficeRuntime.auth.getAccessToken`. This requires
the manifest to declare `ReadWriteDocument` permissions (already set) and the
tenant admin to have not blocked SSO for add-ins (default = allowed). If your
tenant requires explicit consent, the admin needs to consent once via the
integrated apps centre when uploading the manifest.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Ribbon tab doesn't appear | Add-in still deploying, or workbook opened in desktop Excel | Wait ~10 min; use Excel Online. Desktop Excel supports task-pane add-ins too but ribbon contributions can lag. |
| "script_read_url not configured" | `_FlowConfig` placeholder still in place | Paste the real URL into B3 of `_FlowConfig`. |
| 401 on script call | SSO token issue | Sign out of Excel Online, sign back in. Confirm tenant admin allows add-in SSO. |
| 403/404 on flow3_url | Flow 3 not deployed, or URL wrong | Open Flow 3, copy the trigger URL again. |
