# CourtPASS scraper and sync

These scripts collect cases from CourtPASS and send them to the review sheet (see `src/Code.gs`).

| File | What it does |
|---|---|
| `courtpass_scraper.py` | Downloads cases and PDFs from CourtPASS for a date range and writes `index.json`. Retries failed cases, skips stuck PDFs, and saves progress after every case so a rerun picks up where it stopped. |
| `sheet_sync.py` | Reads `index.json` and sends one row per case to the sheet's Apps Script web app. **Issues (CourtPASS)** is sent exactly as scraped. |
| `run_backlog.sh` | Runs the scraper and sync month by month from January 2025 through August 2026. Starts and restarts the scraper's Chrome window as needed. |

## Requirements

macOS, Google Chrome, and Python 3 with `pip3 install playwright python-docx`. CourtPASS blocks headless browsers, so the scraper drives a real Chrome window through its debugging port (9222).

## Backlog

```bash
bash run_backlog.sh scrape    # download every month; rerun to retry failures
bash run_backlog.sh preview   # one-line summary per case in backlog_preview.txt
export SHEET_WEBAPP_URL="https://script.google.com/macros/s/.../exec"
export SHEET_SYNC_TOKEN="same value as SYNC_TOKEN in Script Properties"
bash run_backlog.sh send      # send every month to the sheet; safe to rerun
```

Output goes to `~/courtpass_backlog/<year-month>/`. Never commit the token or the web app URL.

## Single run

```bash
python3 courtpass_scraper.py --start 10/01/2025 --end 10/31/2025 --output ~/courtpass_backlog/2025-10
python3 sheet_sync.py ~/courtpass_backlog/2025-10/index.json --all --dry-run   # preview
python3 sheet_sync.py ~/courtpass_backlog/2025-10/index.json --all             # send
```

## Known issue

**APL no.** is unreliable. CourtPASS removes cases from its docket once decided, so decided cases come back blank, and the docket lookup can pick a different case with similar party names. A fix that uses the court's weekly lists of new appeals is planned.
