#!/bin/bash
# Runs the CourtPASS backlog one month at a time, oldest first.
# Keep this file in the same folder as courtpass_scraper.py and sheet_sync.py.
#
#   bash run_backlog.sh scrape    download every month (safe to rerun: finished months are skipped)
#   bash run_backlog.sh preview   write a one-line summary of every case to backlog_preview.txt
#   bash run_backlog.sh send      send every scraped month to the Sheet (safe to rerun)
#
# Scraping needs the debug Chrome window open. Sending needs SHEET_WEBAPP_URL
# and SHEET_SYNC_TOKEN set in this Terminal window.

cd "$(dirname "$0")" || exit 1
BACKLOG_DIR="${BACKLOG_DIR:-$HOME/courtpass_backlog}"

MONTHS="
2025-01 01/01/2025 01/31/2025
2025-02 02/01/2025 02/28/2025
2025-03 03/01/2025 03/31/2025
2025-04 04/01/2025 04/30/2025
2025-05 05/01/2025 05/31/2025
2025-06 06/01/2025 06/30/2025
2025-07 07/01/2025 07/31/2025
2025-08 08/01/2025 08/31/2025
2025-09 09/01/2025 09/30/2025
2025-10 10/01/2025 10/31/2025
2025-11 11/01/2025 11/30/2025
2025-12 12/01/2025 12/31/2025
2026-01 01/01/2026 01/31/2026
2026-02 02/01/2026 02/28/2026
2026-03 03/01/2026 03/31/2026
2026-04 04/01/2026 04/30/2026
2026-05 05/01/2026 05/31/2026
2026-06 06/01/2026 06/30/2026
2026-07 07/01/2026 07/31/2026
2026-08 08/01/2026 08/31/2026
"

DEBUG_PROFILE="/tmp/chrome_debug"
MONTH_ATTEMPTS=3

chrome_ready() {
  curl -s --max-time 5 http://localhost:9222/json/version >/dev/null 2>&1
}

# Opens the scraper's own Chrome window (separate from your normal Chrome).
start_chrome() {
  echo "--- Starting the scraper's Chrome window"
  open -na "Google Chrome" --args --remote-debugging-port=9222 --user-data-dir="$DEBUG_PROFILE" \
    --disable-background-timer-throttling --disable-renderer-backgrounding --disable-backgrounding-occluded-windows
  for _ in $(seq 1 30); do
    chrome_ready && return 0
    sleep 1
  done
  echo "Chrome didn't start. Open it with the command in README, then run this again."
  return 1
}

# Closes only the scraper's Chrome window, never your normal Chrome.
restart_chrome() {
  echo "--- Restarting the scraper's Chrome window"
  pkill -f "user-data-dir=$DEBUG_PROFILE" 2>/dev/null
  sleep 5
  start_chrome
}

scrape() {
  KEEP_AWAKE=""
  command -v caffeinate >/dev/null 2>&1 && KEEP_AWAKE="caffeinate -i"
  INCOMPLETE=""
  while read -r MONTH START END; do
    [ -z "$MONTH" ] && continue
    OUT="$BACKLOG_DIR/$MONTH"
    if [ -f "$OUT/index.json" ] && [ ! -f "$OUT/progress.json" ]; then
      echo "=== $MONTH: already finished, skipping"
      continue
    fi
    echo "=== $MONTH: scraping"
    mkdir -p "$OUT"
    DONE=""
    for ATTEMPT in $(seq 1 $MONTH_ATTEMPTS); do
      if ! chrome_ready; then
        start_chrome || exit 1
      fi
      echo "--- $MONTH attempt $ATTEMPT, $(date '+%H:%M')" >> "$OUT/run.log"
      $KEEP_AWAKE python3 courtpass_scraper.py --start "$START" --end "$END" --output "$OUT" </dev/null 2>&1 | tee -a "$OUT/run.log"
      if [ "${PIPESTATUS[0]}" -eq 0 ]; then
        DONE=1
        break
      fi
      echo "--- $MONTH stopped (attempt $ATTEMPT of $MONTH_ATTEMPTS). Finished cases are saved."
      [ "$ATTEMPT" -lt "$MONTH_ATTEMPTS" ] && { restart_chrome || exit 1; }
    done
    if [ -z "$DONE" ]; then
      echo "Stopped at $MONTH after $MONTH_ATTEMPTS attempts. Paste the lines above to Claude, then run this command again."
      exit 1
    fi
    [ -f "$OUT/progress.json" ] && INCOMPLETE="$INCOMPLETE $MONTH"
  done <<< "$MONTHS"
  echo
  if [ -n "$INCOMPLETE" ]; then
    echo "Some cases or PDFs failed in:$INCOMPLETE"
    echo "Run this command again to retry only those."
  else
    echo "Scraping finished. Next: bash run_backlog.sh preview"
  fi
}

preview() {
  OUTPUT="backlog_preview.txt"
  : > "$OUTPUT"
  while read -r MONTH START END; do
    [ -z "$MONTH" ] && continue
    INDEX="$BACKLOG_DIR/$MONTH/index.json"
    [ -f "$INDEX" ] || continue
    echo "=== $MONTH" >> "$OUTPUT"
    python3 sheet_sync.py "$INDEX" --all --dry-run --no-url-lookup </dev/null 2>> "$OUTPUT" | python3 -c "
import json, sys
for r in json.load(sys.stdin):
    print(r['session'] or 'NO SESSION', '|', r['decision_date'] or 'NO DECISION', '|', r['title'], '|', r['issues_courtpass'][:70].replace(chr(10), ' / '))" >> "$OUTPUT"
  done <<< "$MONTHS"
  echo "Wrote $OUTPUT"
}

send() {
  if [ -z "$SHEET_WEBAPP_URL" ] || [ -z "$SHEET_SYNC_TOKEN" ]; then
    echo "Set SHEET_WEBAPP_URL and SHEET_SYNC_TOKEN in this Terminal window first."
    exit 1
  fi
  while read -r MONTH START END; do
    [ -z "$MONTH" ] && continue
    INDEX="$BACKLOG_DIR/$MONTH/index.json"
    [ -f "$INDEX" ] || continue
    echo "=== $MONTH: sending"
    if ! python3 sheet_sync.py "$INDEX" --all </dev/null; then
      echo "Stopped at $MONTH. Fix the problem above, then run this command again."
      exit 1
    fi
  done <<< "$MONTHS"
  echo "All scraped months sent."
}

case "$1" in
  scrape) scrape ;;
  preview) preview ;;
  send) send ;;
  *) echo "Use: bash run_backlog.sh scrape | preview | send"; exit 1 ;;
esac
