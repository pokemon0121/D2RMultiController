# D2R Multi Controller — Orchestrator / Worker
**Shortcut-only + Strong PID Map + Post-Launch + Handle Close (Admin-aware)**

- 关闭 *“DiabloII Check For Other Instances”* 依赖 `Utility\handle64.exe`，**需要管理员权限**。
- 我已在代码里：
  - 自动使用 `-accepteula -nobanner`（不会弹 EULA）。
  - 启动前检测管理员权限；若非管理员，`/close_handle` 和 `/launch` 的自动关句柄会返回 `AdminRequired`。
  - 提供 `GET /admin_status` 便于自检。

## 快速启动（Worker）
1) 把 `handle64.exe` 放到 `worker\Utility\handle64.exe`。
2) 以管理员运行：
   - 右键 `RunWorkerAsAdmin.ps1` → 以 PowerShell 运行（会弹 UAC 确认）；或
   - 右键 `RunWorkerAsAdmin.bat` → 以管理员身份运行；或
   - 给 `python.exe`/你的 IDE 设置“以管理员身份运行”。
3) Orchestrator 正常使用。

## 自检
```powershell
curl http://127.0.0.1:5001/admin_status   # -> {"is_admin": true/false}
```

## 注意
- 如果不是管理员，依然可以 **Launch / Join / Layout**，只是**关句柄**这一步会被跳过并给出明确信息。
