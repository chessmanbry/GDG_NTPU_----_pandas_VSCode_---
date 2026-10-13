#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
printf '%s\n' '=====================================================' 'GDG NTPU pandas environment installer - macOS' '====================================================='
if ! command -v python3 >/dev/null 2>&1; then
  echo '[ERROR] Python 3 was not found. Install Python 3.10+ from https://www.python.org/downloads/'
  read -r -p 'Press Enter to close ... ' _
  exit 1
fi
if ! python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1; then
  echo '[ERROR] This workshop requires Python 3.10 or later.'
  read -r -p 'Press Enter to close ... ' _
  exit 1
fi
echo '[1/3] Creating local virtual environment .venv ...'
python3 -m venv .venv
echo '[2/3] Updating pip ...'
".venv/bin/python" -m pip install --upgrade pip
echo '[3/3] Installing pandas and openpyxl ...'
".venv/bin/python" -m pip install -r requirements.txt
".venv/bin/python" -c 'import pandas, openpyxl; print("READY: pandas", pandas.__version__, "openpyxl", openpyxl.__version__)'
echo '[PASS] Setup completed. Next: run_mac.command'
read -r -p 'Press Enter to close ... ' _
