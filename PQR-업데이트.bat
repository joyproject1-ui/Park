@echo off
rem ============================================================
rem  PQR program updater (Windows)
rem
rem  ASCII only on purpose. A batch file that mixes non-ASCII text
rem  with cmd.exe code pages garbles every line after it, so all
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
  echo  Install it from https://www.python.org/downloads/ first.
  echo.
  pause
  exit /b 1
)

rem Install or refresh the libraries the auto-report engine needs.
rem Output goes to "설치 기록.txt" so a failed install (e.g. the handwriting reader) can be diagnosed.
%PY% -m pip install -q -r requirements.txt --disable-pip-version-check > "설치 기록.txt" 2>&1
if errorlevel 1 (
  echo.
  echo   [!] Some libraries failed to install. See "설치 기록.txt" and send it to the maintainer.
  echo       The report still works; the handwriting reader for item 13 may be off.
  echo.
)
%PY% -m pqr update
pause
