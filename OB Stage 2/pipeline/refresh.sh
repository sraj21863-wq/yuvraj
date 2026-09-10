#!/bin/sh
# Rebuild the offline snapshot (dashboard/data.js) from the workbook.
cd "$(dirname "$0")" || exit 1
python3 -c 'import openpyxl' 2>/dev/null || python3 -m pip install -r requirements.txt
exec python3 ob_pipeline.py
