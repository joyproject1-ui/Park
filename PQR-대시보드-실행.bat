@echo off
rem ============================================================
rem  PQR dashboard launcher (Windows)
rem
rem  This file is deliberately ASCII-only. A batch file that mixes
rem  "chcp" with non-ASCII text makes cmd.exe lose its place while
rem  reading the script, which garbles every line after it. All
rem  Korean messages are printed by Python instead.
rem ============================================================

cd /d "%~dp0"

set "CHECK=import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)"
set "PY="

py -3 -c "%CHECK%" >nul 2>&1 && set "PY=py -3"
if not defined PY (
  python -c "%CHECK%" >nul 2>&1 && set "PY=python"
)
if not defined PY (
  python3 -c "%CHECK%" >nul 2>&1 && set "PY=python3"
)

if not defined PY (
  echo.
  echo  [ERROR] Python 3.9 or newer was not found.
  echo.
  echo   1^) Download it from https://www.python.org/downloads/
  echo   2^) Tick "Add Python to PATH" on the first setup screen.
  echo   3^) Run this file again.
  echo.
  echo   Note: if a Microsoft Store window opened, that is not a real
  echo         installation - use the link above.
  echo.
  pause
  exit /b 1
)

rem Install or refresh the libraries the auto-report engine needs.
rem Output goes to install-log.txt so a failed install (e.g. the handwriting reader) can be diagnosed.
%PY% -m pip install -q -r requirements.txt --disable-pip-version-check > install-log.txt 2>&1
if errorlevel 1 (
  echo.
  echo   [!] Some core libraries failed to install. See install-log.txt and send it to the maintainer.
  echo.
)
rem Handwriting reader for item 13 (installed separately; the report still works without it).
rem rapidocr is the reader, onnxruntime its engine. Python 3.13 needs this pair; the old
rem package name rapidocr-onnxruntime stops at 3.12.
%PY% -c "import onnxruntime, rapidocr" >nul 2>&1
if errorlevel 1 (
  %PY% -m pip install -q -r requirements-ocr.txt --disable-pip-version-check > install-ocr-log.txt 2>&1
  if errorlevel 1 (
    echo.
    echo   [!] The handwriting reader could not be installed. See install-ocr-log.txt and send it.
    echo       The report still works; item 13 will carry last year's values with a yellow note.
    echo.
  )
)
%PY% -m pqr launch
if errorlevel 1 pause
