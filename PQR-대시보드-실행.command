#!/bin/bash
# PQR 대시보드 실행 (macOS · Linux)
cd "$(dirname "$0")" || exit 1

INPUT="입력폴더"
PORT=8787

PY=$(command -v python3 || command -v python) || {
  echo "파이썬을 찾지 못했습니다. python.org 에서 3.9 이상을 설치하세요."
  read -r -p "엔터를 누르면 닫힙니다."
  exit 1
}

mkdir -p "$INPUT/공통"
echo "브라우저에서 열기:  http://127.0.0.1:$PORT"
echo "종료하려면 Ctrl+C 를 누르세요."
( sleep 2; (command -v open >/dev/null && open "http://127.0.0.1:$PORT") || \
  (command -v xdg-open >/dev/null && xdg-open "http://127.0.0.1:$PORT") ) &
exec "$PY" -m pqr serve --in "$INPUT" --port "$PORT"
