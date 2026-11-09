#!/bin/bash
cd "$(dirname "$0")"
if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display dialog "Python 3 is required. Install it from python.org, then open this again." buttons {"OK"}'
  exit 1
fi
python3 act10_app.py
