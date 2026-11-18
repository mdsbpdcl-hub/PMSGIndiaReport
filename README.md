# PM Surya Ghar – Bihar in India

Daily state-wise progress of **PM Surya Ghar: Muft Bijli Yojana**, with Bihar's position.

| Part | File | What it does |
|---|---|---|
| Dashboard | `index.html` | Bihar's rank, installations, ranking table, ranking among free-power states, automatic findings, day-by-day trend. Reads the live figures from the "PMSG-MBY All States" sheet. |
| Portal reader | `scraper/pmsg_portal_scraper.py` | Opens pmsuryaghar.gov.in (State/UT-wise + Month-wise), saves Excel/CSV and pushes the figures to the sheet. |
| One-off copy | `scraper/pmsg_portal_console.js` | Paste in the browser console on the portal page → downloads a CSV. |
| Sheet updater | `apps-script/PMSG_Sheet_AutoUpdate.gs` | Sorts by installations, ranks, adds free-power column, totals, change since last update, History tab. Backup daily update if no PC push. |
| Daily run | `.github/workflows/daily-update.yml` | Runs the portal reader every day at about 09:11 IST on the office PC. |

> The portal blocks servers outside India, so the daily run must happen on a PC in India (self-hosted runner). GitHub's own servers cannot read it.

## Setup

### 1. Sheet
1. Open the sheet → **Extensions → Apps Script** → paste `apps-script/PMSG_Sheet_AutoUpdate.gs` → change `TOKEN` → Save.
2. Run `installDailyTrigger` once and allow permissions.
3. **Deploy → New deployment → Web app** (Execute as: Me, Who has access: Anyone) → copy the URL.
4. **Share → Anyone with the link → Viewer** (needed for the dashboard to read the figures).

### 2. GitHub secrets
Repository **Settings → Secrets and variables → Actions → New repository secret**:
- `SHEET_WEBAPP_URL` = the web-app URL from step 1.3
- `SHEET_TOKEN` = the same secret word as `TOKEN`

### 3. Office PC as runner (Windows, once)
1. Install **Python 3.11+** (tick "Add to PATH") and **Git**.
2. Repository **Settings → Actions → Runners → New self-hosted runner → Windows**; run the commands shown, then `./svc.cmd install` and `./svc.cmd start` so it runs as a service.
3. **Settings → Actions → General → Fork pull request workflows**: keep "Require approval for all external contributors".
4. **Actions → Daily PM Surya Ghar update → Run workflow** to test.

The PC must be on and connected around 09:11 IST. If it is off, the sheet updates itself at 10:00 IST from a public daily copy of the portal.

### 4. Dashboard on GitHub Pages
**Settings → Pages → Deploy from a branch → `main` / root → Save.**
The dashboard will be at `https://mdsbpdcl-hub.github.io/PMSGIndiaReport/`.

To change the free-electricity notes, edit `FREE_POWER` in the Apps Script and in the Python script.
