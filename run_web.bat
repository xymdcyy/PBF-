@echo off
setlocal
cd /d "%~dp0"
title PBF Web UI

echo ==========================================
echo   PBF Rolling Forecast - Web UI
echo ==========================================
echo.

REM Locate the project venv my_finance_env (try a few relative paths).
set "PY="
if exist "%~dp0my_finance_env\Scripts\python.exe"   set "PY=%~dp0my_finance_env\Scripts\python.exe"
if not defined PY if exist "%~dp0..\my_finance_env\Scripts\python.exe"     set "PY=%~dp0..\my_finance_env\Scripts\python.exe"
if not defined PY if exist "%~dp0..\..\my_finance_env\Scripts\python.exe"  set "PY=%~dp0..\..\my_finance_env\Scripts\python.exe"

if defined PY goto have_py

REM No venv found: fall back to python on PATH.
set "PY=python"
where python >nul 2>nul
if errorlevel 1 goto no_python

:have_py
echo [ENV] Python: %PY%
echo [RUN] %PY% -m webapp.app
echo.
echo Open in browser: http://127.0.0.1:8765
echo Press Ctrl+C to quit.
echo.

start "" http://127.0.0.1:8765
"%PY%" -m webapp.app
exit /b %ERRORLEVEL%

:no_python
echo [ERROR] Python not found.
echo         Install Python 3.10+ and run:  pip install -r requirements.txt
pause
exit /b 1
