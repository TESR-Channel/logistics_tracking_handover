# TESR — Tracking & Handover Record

Single-file web app for TESR Shop delivery handover records, deployed on GitHub Pages.

- `index.html` — the whole app (UI, customer database, logo). Nothing to build.
- `tesr-drive-backend.gs` — Google Apps Script that saves each record to Google Drive
  (`TESR shop tracking and handover / <year> / <Tax ID> / INVxxxxx_DDMMYYYY`) and logs to a Google Sheet.

## Deploy on GitHub Pages
1. Push this folder to a repository (files at the root).
2. Repo → **Settings → Pages → Build and deployment**: Source = *Deploy from a branch*, Branch = `main`, Folder = `/ (root)` → Save.
3. Open `https://tesr-channel.github.io/logistics_tracking_handover/` (takes 1–2 min the first time).

## Connect Google Drive
1. https://script.google.com → New project → paste `tesr-drive-backend.gs`.
2. Deploy → New deployment → Web app · Execute as **Me** · Who has access **Anyone** → copy the `/exec` URL.
3. In the app: Settings → Live mode → paste URL → Save settings.

Records/Tracking List are stored in each browser (localStorage). The shared source of truth is the Google Sheet log the backend creates.
