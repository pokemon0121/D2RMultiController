:: run_worker_admin.bat
@echo off
REM Start D2R Worker (auto request admin)

:: always run from script folder
cd /d %~dp0

:: request admin if needed
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: prefer venv python
set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python.exe"

:: fixed config in project root + default port
set "CONFIG=%CD%\config.json"
set "PORT=5001"

:: map COMPUTERNAME -> worker name (edit these lines only)
set "WORKER_NAME="
if /I "%COMPUTERNAME%"=="RGB"       set "WORKER_NAME=Worker-MSI-Desktop"
if /I "%COMPUTERNAME%"=="ROC-YIMU"  set "WORKER_NAME=Worker-ASUS-ROG-Laptop"

if "%WORKER_NAME%"=="" (
    echo [worker] ERROR: no mapping for COMPUTERNAME=%COMPUTERNAME%. Edit this BAT.
    pause
    exit /b 1
)

:: launch
"%PY%" worker\worker.py --name "%WORKER_NAME%" --config "%CONFIG%" --port %PORT%
