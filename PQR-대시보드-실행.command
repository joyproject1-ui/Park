#!/bin/bash
# PQR 대시보드 실행 (macOS · Linux)
cd "$(dirname "$0")" || exit 1

INPUT="입력폴더"
PORT=8787

CHECK='import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)'
PY=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c "$CHECK" >/dev/null 2>&1; then
    PY="$candidate"; break
  fi
done
if [ -z "$PY" ]; then
  echo "파이썬 3.9 이상을 찾지 못했습니다. python.org 에서 설치한 뒤 다시 실행하세요."
  read -r -p "엔터를 누르면 닫힙니다."
  exit 1
fi
echo "파이썬 확인: $("$PY" --version 2>&1)"

mkdir -p "$INPUT/공통"
echo "브라우저에서 열기:  http://127.0.0.1:$PORT"
echo "종료하려면 Ctrl+C 를 누르세요."
( sleep 2; (command -v open >/dev/null && open "http://127.0.0.1:$PORT") || \
  (command -v xdg-open >/dev/null && xdg-open "http://127.0.0.1:$PORT") ) &
exec "$PY" -m pqr serve --in "$INPUT" --port "$PORT"
