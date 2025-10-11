@echo off
REM run_orchestrator.bat —— 启动 Tk UI（建议配合快捷方式“最小化运行”）
setlocal
pushd "%~dp0"

set "PYW=.venv\Scripts\pythonw.exe"
if not exist "%PYW%" set "PYW=pythonw.exe"

REM 显式指定根目录的 config.json
set "CMD=%PYW% orchestrator\orchestrator_ui.py --config ""%CD%\config.json"""

REM 方式 A：后台启动（/b），窗口最小化（/min）
start "" /min /b %CMD%

popd
endlocal
