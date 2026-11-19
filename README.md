# PM Surya Ghar – Bihar in India

Daily state-wise progress of **PM Surya Ghar: Muft Bijli Yojana**, with Bihar's position.

| Part | File | What it does |
|---|---|---|
| Dashboard | `index.html` (+ `vendor/` PDF library) | Export PDF of free-electricity states and all states. Bihar's rank, installations, ranking table, ranking among free-power states, automatic findings, day-by-day trend. Reads the live figures from the "PMSG-MBY All States" sheet. |
| Portal reader | `scraper/pmsg_portal_scraper.py` | Opens pmsuryaghar.gov.in (State/UT-wise + Month-wise), saves Excel/CSV and pushes the figures to the sheet. |
| One-off copy | `scraper/pmsg_portal_console.js` | Paste in the browser console on the portal page → downloads a CSV. |
| Sheet updater | `apps-script/PMSG_Sheet_AutoUpdate.gs` | Sorts by installations, ranks, adds free-power column, totals, change since last update, History tab. Backup daily update if no PC push. |
| Daily run | `.github/workflows/daily-update.yml` | Runs the portal reader every day at about 09:11 IST on the office PC. |

## How it updates (fully online)

1. **9 AM and 6 PM daily** – the script inside the "PMSG-MBY All States" sheet reads the daily data file of the public copy of the National Portal (solentra.in/data/pmsg-snapshot.json, unofficial), sorts states by installations, adds ranks, free-electricity details, India total, change since last update, and a History tab.
2. **Dashboard** – reads the sheet every time it is opened or refreshed.

No PC is needed. If the copy site fails or changes, the sheet keeps the last good figures (the dashboard shows their date).
The portal itself cannot be read from Google or GitHub servers (it blocks servers outside India).

## Setup (about 10 minutes, once)

1. **Sheet script** – open the sheet → **Extensions → Apps Script** → delete any code → paste `apps-script/PMSG_Sheet_AutoUpdate.gs` → **Save**.
2. Pick **installDailyTrigger** in the function list → **Run** → allow the permissions. The State-wise tab fills immediately.
3. **Share the sheet**: **Share → General access → Anyone with the link → Viewer** (the dashboard reads it).
4. **Dashboard online**: this repository → **Settings → Pages → Deploy from a branch → `main` / root → Save**.
   Dashboard: `https://mdsbpdcl-hub.github.io/PMSGIndiaReport/`

5. **Force update button on the dashboard** – in Apps Script set `TOKEN` (line near the top) to your password, then **Deploy → New deployment → Select type: Web app → Execute as: Me, Who has access: Anyone → Deploy**, and **open the Web app link once** (it saves the link for the dashboard).
   Keep the password only in Apps Script – never commit it to this repository.

Manual refresh any time: dashboard **Force update** (password) → *Update now*, or in the sheet **PMSG Update → Update now**.
Official portal data any time: dashboard **Force update** → *Download program* → run it on a computer in India.

## Optional: official portal data from an office PC
The Python reader (`scraper/`) can run on a PC in India and push the official portal figures to the sheet. Deploy the Apps Script as a **Web app** (Execute as: Me, Access: Anyone), set `TOKEN`, add GitHub secrets `SHEET_WEBAPP_URL` and `SHEET_TOKEN`, register the PC as a **self-hosted runner** (Settings → Actions → Runners), and switch on the schedule in `.github/workflows/daily-update.yml`.

To change the free-electricity notes, edit `FREE_INFO` in `index.html` and `FREE_POWER` in the Apps Script.
