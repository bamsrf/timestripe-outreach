#!/bin/bash
# Daily auto-send: pull replies first, then send N welcome emails.
# Run by launchd (see com.timestripe.daily-welcome.plist).
#
# Logs go to output/daily_welcome.log (append, with timestamp).

set -u
cd "$(dirname "$0")/.." || exit 1

LIMIT="${1:-25}"
LOG=output/daily_welcome.log
PY=.venv/bin/python

mkdir -p output

{
  echo ""
  echo "======================================================================"
  echo "=== $(date '+%Y-%m-%d %H:%M:%S') === Daily auto-send (limit=$LIMIT)"
  echo "======================================================================"

  # Sanity: env + venv must exist
  if [ ! -x "$PY" ]; then
    echo "ERROR: venv not found at $PY. Set up with: python3 -m venv .venv && pip install -r requirements.txt"
    exit 1
  fi
  if [ ! -f .env ]; then
    echo "ERROR: .env missing"
    exit 1
  fi

  # 1) Pull fresh replies from IMAP — so we don't send to people who already replied
  echo ""
  echo "--- Step 1/2: syncing replies from inbox ---"
  $PY reply_detector.py --apply --since 2 2>&1 || echo "! reply sync failed but proceeding"

  # 2) Send the welcome batch
  echo ""
  echo "--- Step 2/2: sending $LIMIT welcome emails ---"
  $PY outreach_sender.py --send --limit "$LIMIT" 2>&1
  EXIT=$?

  echo ""
  echo "=== Done. Sender exit code: $EXIT ==="
  exit $EXIT
} >> "$LOG" 2>&1
