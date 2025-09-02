:: run_worker_admin.bat
@echo off
REM 启动 D2R Worker (自动请求管理员权限)

:: 切换到脚本所在目录，避免在 System32 下找不到文件
cd /d %~dp0

:: 检查是否有管理员权限
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: 默认使用虚拟环境 python
.venv\Scripts\python.exe worker\worker.py --config worker\config.json

pause