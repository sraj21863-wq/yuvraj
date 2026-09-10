#!/bin/sh
# Start the Stage 2 data service (macOS / Linux).
cd "$(dirname "$0")" || exit 1
python3 -c 'import openpyxl' 2>/dev/null || python3 -m pip install -r requirements.txt
exec python3 serve.py
