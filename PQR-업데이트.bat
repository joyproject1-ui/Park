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
rem Output goes to install-log.txt so a failed install (e.g. the handwriting reader) can be diagnosed.
rem (ASCII only in this file: cmd.exe reads .bat files in the OEM code page and breaks on Korean.)
%PY% -m pip install -q -r requirements.txt --disable-pip-version-check > install-log.txt 2>&1
if errorlevel 1 (
  echo.
  echo   [!] Some libraries failed to install. See install-log.txt and send it to the maintainer.
  echo       The report still works; the handwriting reader for item 13 may be off.
  echo.
)
rem The updater overwrites this very file. cmd.exe keeps reading a running .bat by byte
rem offset, so anything after the update on later lines would be read from the NEW file at
rem the OLD offset ('pdater', 'y' errors). Keep update, pause and exit on ONE line.
%PY% -m pqr update & pause & exit /b 0
