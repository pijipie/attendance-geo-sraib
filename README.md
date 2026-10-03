# attendance-geo-sraib

Geo-fenced staff attendance website. Static pages on GitHub Pages, backed by a Google Apps Script web app that reads and writes a Google Sheet.

## Live pages

- [v3.html](https://pijipie.github.io/attendance-geo-sraib/v3.html) — staff check-in (newest)
- [v3-design.html](https://pijipie.github.io/attendance-geo-sraib/v3-design.html) — redesigned check-in (same backend as v3)
- [admin.html](https://pijipie.github.io/attendance-geo-sraib/admin.html) — PIN-protected daily attendance report
- [v2.html](https://pijipie.github.io/attendance-geo-sraib/v2.html), [index.html](https://pijipie.github.io/attendance-geo-sraib/index.html) — older versions

## Backend

`apps-script/Code.gs` is the Apps Script project bound to the Google Sheet. It expects these tabs: `Config`, `Staff`, `Events`, `Records`, and (optional) `Summary`.

After editing `Code.gs`, paste it into the Apps Script editor and **Deploy → Manage deployments → Edit → New version** so the existing `/exec` URL keeps working.

> **Do NOT commit staff lists, attendance exports, venue coordinates or the admin PIN to this public repo.** They live in the Google Sheet only.
