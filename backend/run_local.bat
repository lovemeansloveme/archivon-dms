@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
  echo [Dowple] 가상환경이 없습니다.
  echo 먼저 아래 명령을 실행해주세요.
  echo.
  echo cd backend
  echo python -m venv .venv
  echo .venv\Scripts\activate.bat
  echo python -m pip install -r requirements.txt
  echo.
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"

echo [Dowple] FastAPI 서버를 시작합니다.
echo [Dowple] http://127.0.0.1:8001/health 에서 확인할 수 있습니다.
python -m uvicorn main:app --host 127.0.0.1 --port 8001

endlocal
