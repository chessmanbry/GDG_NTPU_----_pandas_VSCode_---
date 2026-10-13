#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
if [ ! -f '.venv/bin/python' ]; then
  echo '[INFO] Local dependencies missing; starting setup_mac.command ...'
  bash './setup_mac.command'
fi
".venv/bin/python" run_audit.py
".venv/bin/python" verify_results.py
echo '[PASS] Results are saved in outputs/'
read -r -p 'Press Enter to close ... ' _
