@echo off
cd /d "%~dp0\.."
if not "%ROTA_USE_EXTERNAL_ENV%"=="1" set "ROTA_DB=%CD%\rotadb.db"
if "%ROTA_DB%"=="" set "ROTA_DB=%CD%\rotadb.db"
if not "%ROTA_USE_EXTERNAL_ENV%"=="1" set "DATABASE_URL=sqlite+aiosqlite:///%ROTA_DB:\=/%"
if "%DATABASE_URL%"=="" set "DATABASE_URL=sqlite+aiosqlite:///%ROTA_DB:\=/%"
if not "%ROTA_USE_EXTERNAL_ENV%"=="1" set "ROTA_MOBILE_PHOTOS_DIR=%CD%\.rotahub_runtime\fotos_rotas"
if "%ROTA_MOBILE_PHOTOS_DIR%"=="" set "ROTA_MOBILE_PHOTOS_DIR=%CD%\.rotahub_runtime\fotos_rotas"
if "%ROTA_ENABLE_LEGACY_MOBILE_API%"=="" set "ROTA_ENABLE_LEGACY_MOBILE_API=1"
if "%HOST%"=="" set "HOST=0.0.0.0"
if "%PORT%"=="" set "PORT=8000"
if "%PYTHON_EXE%"=="" set "PYTHON_EXE=python"
"%PYTHON_EXE%" -m uvicorn backend.main:app --host %HOST% --port %PORT%
