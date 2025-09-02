' run_orchestrator.vbs —— 无控制台启动 Tk UI
Option Explicit
Dim fso, shell, here
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
here = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = here

' 用 pythonw.exe 启动 UI，WindowStyle=0（隐藏），bWaitOnReturn=False（后台运行）
shell.Run ".venv\Scripts\pythonw.exe orchestrator\orchestrator_ui.py --config orchestrator\config.json", 0, False