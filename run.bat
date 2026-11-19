@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title PBF Rolling Forecast Automation

echo ==========================================
echo   PBF Rolling Forecast Automation
echo ==========================================
echo   导出多期间数据: export.bat --periods 2027.01-2027.04
echo   批量清除填 0 : clear.bat  --periods 2027.01-2027.04 [--run]
echo.

set "RUNNER=pbf-uploader.exe"
if exist pbf-uploader.exe goto run
where python >nul 2>nul
if errorlevel 1 goto no_python
set "RUNNER=python cli.py"

:run
echo [RUN] %RUNNER% run-flow %*
echo.
%RUNNER% run-flow %*
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" echo [DONE] Flow finished. See the indicator table above.
if not "%RC%"=="0" echo [FAILED] Exit code %RC%. Resume with: run.bat --from-step N
echo.
pause
exit /b %RC%

:no_python
echo [ERROR] Neither pbf-uploader.exe nor python was found.
echo         Put pbf-uploader.exe next to this script, or install Python 3.10+.
echo.
pause
exit /b 1
