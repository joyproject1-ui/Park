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
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
  python --version >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo.
  echo  [문제] 파이썬을 찾지 못했습니다.
  echo.
  echo  https://www.python.org/downloads/ 에서 파이썬 3.9 이상을 설치하세요.
  echo  설치 화면에서 "Add Python to PATH" 를 반드시 체크해야 합니다.
  echo  설치 후 이 파일을 다시 실행하세요.
  echo.
  pause
  exit /b 1
)

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
