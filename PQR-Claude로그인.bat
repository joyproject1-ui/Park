@echo off
rem ============================================================
rem  Log in to Claude Code on this PC (Windows)
rem
rem  The installer window closes when it ends, so `claude` cannot be
rem  typed there. This file opens Claude and keeps the window open.
rem  ASCII only on purpose (see PQR-jindan.bat comment).
rem ============================================================

cd /d "%~dp0"
title Claude login

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

%PY% -m pqr login-claude
echo.
pause
