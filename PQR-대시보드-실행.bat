@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  PQR 대시보드 실행 (Windows)
rem  이 파일을 더블클릭하면 대시보드가 열립니다.
rem ============================================================

cd /d "%~dp0"

set "INPUT=입력폴더"
set "PORT=8787"

echo.
echo  PQR 대시보드를 시작합니다.
echo  ------------------------------------------------------------

rem --- 파이썬 찾기 ---
rem  단순히 명령이 있는지만 보면, 윈도우에 기본으로 들어 있는 Microsoft Store 연결용
rem  껍데기(python.exe)를 진짜 설치로 착각합니다. 실제로 3.9 이상이 실행되는지 확인합니다.
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
  echo  [문제] 파이썬 3.9 이상을 찾지 못했습니다.
  echo.
  echo  1) https://www.python.org/downloads/ 에서 내려받아 설치하세요.
  echo  2) 설치 첫 화면 아래의 "Add Python to PATH" 를 반드시 체크하세요.
  echo  3) 설치가 끝나면 이 파일을 다시 실행하세요.
  echo.
  echo  * Microsoft Store 창이 열린 적이 있다면, 그것은 실제 설치가 아닙니다.
  echo    위 주소에서 정식 설치본을 받아 설치해 주세요.
  echo.
  pause
  exit /b 1
)

for /f "tokens=*" %%v in ('%PY% --version 2^>^&1') do echo  파이썬 확인: %%v

rem --- 입력 폴더 준비 ---
if not exist "%INPUT%" (
  mkdir "%INPUT%"
  echo  입력 폴더를 만들었습니다: %INPUT%
)
if not exist "%INPUT%\공통" (
  mkdir "%INPUT%\공통"
  echo  공통 폴더를 만들었습니다: %INPUT%\공통
  echo.
  echo  [먼저 할 일] 제품 마스터 파일을 %INPUT%\공통 폴더에 넣어 주세요.
  echo               넣은 뒤 화면 오른쪽 위의 새로고침 단추를 누르면 반영됩니다.
)

rem --- 대시보드 열기 ---
echo.
echo  브라우저에서 열기:  http://127.0.0.1:%PORT%
echo  종료하려면 이 창에서 Ctrl+C 를 누르거나 창을 닫으세요.
echo  ------------------------------------------------------------
echo.
start "" "http://127.0.0.1:%PORT%"

%PY% -m pqr serve --in "%INPUT%" --port %PORT%

if errorlevel 1 (
  echo.
  echo  [문제] 대시보드를 시작하지 못했습니다. 위 메시지를 확인하세요.
  echo  포트 %PORT% 가 이미 쓰이고 있다면 이 파일에서 PORT 값을 8788 로 바꿔 보세요.
  echo.
  pause
)
endlocal
