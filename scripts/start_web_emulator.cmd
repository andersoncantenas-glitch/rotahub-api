@echo off
cd /d "%~dp0\.."
set "DATABASE_URL=sqlite+aiosqlite:///C:/pdc_rota/banco.db"
set "ROTA_DB=C:/pdc_rota/banco.db"
set "APP_ENV=development"
set "ROTA_SECRET=rota-secreta"
set "ALLOWED_HOSTS=localhost,127.0.0.1,0.0.0.0,10.0.2.2"
set "CORS_ORIGINS=http://localhost:8000,http://127.0.0.1:8000,http://10.0.2.2:8000"
set "ROTA_ENABLE_LEGACY_MOBILE_API=1"
set "HOST=0.0.0.0"
set "PORT=8000"
set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"
"%PYTHON_EXE%" -m uvicorn backend.main:app --host %HOST% --port %PORT%
pause
