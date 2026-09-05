@echo off
rem ============================================================
rem  PQR program updater (Windows)
rem
rem  ASCII only on purpose: cmd.exe reads .bat files in the OEM
rem  codepage, and non-ASCII text corrupts the lines that follow.
rem  All Korean messages are printed by Python instead.
rem
rem  The updater replaces this very file. cmd.exe keeps reading a
rem  running .bat by byte offset, so the update must never run from
rem  the file it replaces: copy this file to TEMP and run the copy.
rem ============================================================

if /i not "%~dp0"=="%TEMP%\" (
  copy /y "%~f0" "%TEMP%\PQR-update-run.bat" >nul
  "%TEMP%\PQR-update-run.bat" "%~dp0"
  exit /b
)
set "ROOT=%~1"
if "%ROOT%"=="" set "ROOT=%CD%"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
cd /d "%ROOT%"

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
%PY% -m pip install -q -r requirements.txt --disable-pip-version-check > install-log.txt 2>&1
if errorlevel 1 (
  echo.
  echo   [!] Some libraries failed to install. See install-log.txt and send it to the maintainer.
  echo       The report still works; the handwriting reader for item 13 may be off.
  echo.
)
%PY% -m pqr update --dir "%ROOT%"
pause
exit /b 0
