# attendance-geo-sraib — notes for Claude

Version 1 of the staff attendance system (Google Sheets + Apps Script). Version 2, being rebuilt on Supabase, lives in the sibling repo `attendance-app`.

## Layout
- `v3.html`, `v3-design.html` — staff check-in pages; `admin.html` — admin report. All call the Apps Script `/exec` URL in `URL_APP`.
- `index.html`, `v2.html` — older pages that use `WEB_APP_URL`; keep them working but don't develop them further.
- `apps-script/Code.gs` — backend. Routes on `action` in `doPost`: `verifyStaff`, `preview`, `adminReport`, otherwise check-in.

## Rules
- Public repo, served by GitHub Pages from `main`. Never commit staff names, CSV/XLSX exports, venue coordinates or the AdminPIN.
- Venue coordinates must never be sent to the browser. The server returns only distance and inside/outside.
- UI text is in Bahasa Melayu; keep new messages in Malay.
- Plain HTML/CSS/JS with no build step. Apps Script uses ES5-style `var`/`function`, so match it.
- Changing `Code.gs` requires a manual redeploy in the Apps Script editor (new version on the same deployment).
- Work on a branch and open a PR. Don't push straight to `main`.
