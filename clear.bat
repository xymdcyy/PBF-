@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title PBF Clear Forecast Data

set "RUNNER=pbf-uploader.exe"
if exist pbf-uploader.exe goto have_runner
where python >nul 2>nul
if errorlevel 1 goto no_python
set "RUNNER=python cli.py"
:have_runner

echo ==========================================
echo   PBF 批量清除（把填过的数字清成 0）
echo ==========================================
echo   用法: clear.bat --periods 2027.01-2027.04 [--run]
echo.
%RUNNER% clear %*
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" echo [DONE] Clear finished.
if not "%RC%"=="0" echo [FAILED] Exit code %RC%.
echo.
pause
exit /b %RC%

:no_python
echo [ERROR] Neither pbf-uploader.exe nor python was found.
echo         Put pbf-uploader.exe next to this script, or install Python 3.10+.
echo.
pause
exit /b 1
