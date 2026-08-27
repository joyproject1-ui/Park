#!/bin/bash
# PQR dashboard launcher (macOS / Linux)
cd "$(dirname "$0")" || exit 1

CHECK='import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)'
PY=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c "$CHECK" >/dev/null 2>&1; then
    PY="$candidate"; break
  fi
done

if [ -z "$PY" ]; then
  echo "Python 3.9 or newer was not found. Install it from https://www.python.org/downloads/"
  read -r -p "Press Enter to close."
  exit 1
fi

exec "$PY" -m pqr launch
