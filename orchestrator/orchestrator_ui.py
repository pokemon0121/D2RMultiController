import json
import random
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import font as tkfont
from datetime import datetime
from collections import defaultdict
import tempfile, os
from datetime import timezone
from tkinter import ttk 
from typing import Callable, Dict, Any, Iterable, List, Optional
from functools import partial
import requests

CFG_PATH = (Path(__file__).resolve().parent.parent / "config.json")

COLOR_MAP = {
    "1": "DeepSkyBlue4",
    "2": "DarkGreen",
    "3": "Firebrick3",
    "4": "Purple3",
    "5": "DarkOrange3",
    "6": "DarkCyan",
    "7": "Magenta3",
    "8": "SaddleBrown",
}

_autosave_timer = None
def _atomic_write_json(path: Path, data: dict):
    # overwrite-only (no backup)
    tmp = path.with_suffix(".tmp")
    txt = json.dumps(data, indent=2, ensure_ascii=False)
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(txt)
        f.flush()
        os.fsync(f.fileno())
    tmp.replace(path)


def save_config_debounced():
    global _autosave_timer, cfg
    try:
        cfg["updated_at"] = datetime.now(timezone.utc).astimezone().isoformat()
        cfg["workers"] = WORKERS
        cfg["targets"] = TARGETS
        cfg["assignment"] = ASSIGN
        cfg["prefs"] = PREFS
    except Exception as e:
        log_target(f"[config] compose failed: {e}")
        return

    def _do_save():
        global _autosave_timer
        _autosave_timer = None
        try:
            _atomic_write_json(CFG_PATH, cfg)
            log_target(f"Saved {CFG_PATH.name} @ {datetime.now().strftime('%H:%M:%S')}")
        except Exception as e:
            log_target(f"Save error: write {CFG_PATH.name} failed: {e}")

    delay_ms = int(PREFS.get("autosave_debounce_ms", 400))
    if _autosave_timer:
        root.after_cancel(_autosave_timer)
    _autosave_timer = root.after(delay_ms, _do_save)

def load_config():
    if not CFG_PATH.exists():
        log_target(f"[config] missing file: {CFG_PATH}")
        return None
    return json.loads(CFG_PATH.read_text(encoding="utf-8"))

cfg = load_config()
if cfg is None:
    raise SystemExit(1)

# 新结构：单一 config.json
WORKERS = cfg["workers"]                 # dict[str, {url, enabled, d2r_root, join_delay_sec, ...}]
TARGETS = cfg["targets"]                 # dict[str, {name, lnk}]
ASSIGN  = cfg["assignment"]              # dict[str, worker_name]
PREFS   = cfg.get("prefs", {})           # { autosave_debounce_ms, keep_backups }
UPDATED_AT = cfg.get("updated_at", "")

# 目标运行状态：Idle / Launching / Running / Stopping / Error
STATE: dict[str, str] = {tid: "Idle" for tid in TARGETS.keys()}

# UI 变量与控件引用
assign_vars: dict[str, tk.StringVar] = {}
lnk_labels: dict[str, tk.Label] = {}
worker_cmbs: dict[str, ttk.Combobox] = {}
lnk_entries: dict[str, ttk.Entry] = {}
status_labels: dict[str, tk.Label] = {}

def is_idle(tid: str) -> bool:
    return STATE.get(str(tid), "Idle") == "Idle"

def resolved_shortcut_for(tid: str) -> str | None:
    tid = str(tid)
    wn = ASSIGN.get(tid)
    if not wn:
        return None
    t = TARGETS.get(tid, {}) or {}
    lnk = (t.get("lnk") or "").strip()
    if not lnk:
        return None

    low = lnk.lower()
    if not (low.endswith(".lnk") or low.endswith(".exe") or low.endswith(".bat")):
        lnk = lnk + ".lnk"

    root_dir = (WORKERS.get(wn, {}).get("d2r_root") or "").rstrip("\\/")
    if not root_dir:
        return None
    return root_dir + "\\" + lnk


def refresh_row_enabled(tid: str):
    tid = str(tid)
    idle = is_idle(tid)
    # Name 是 Label，始终只读
    if cmb := worker_cmbs.get(tid):
        cmb.configure(state=("readonly" if idle else "disabled"))
    if lab := status_labels.get(tid):
        lab.configure(text=STATE.get(tid, "Idle"))

def refresh_all_rows():
    for tid in TARGETS.keys():
        refresh_row_enabled(tid)

def _on_assign_change(tid: str, *_):
    tid = str(tid)
    new_wn = assign_vars[tid].get().strip()
    if not new_wn or new_wn not in WORKERS:
        # 回退
        assign_vars[tid].set(ASSIGN.get(tid, ""))
        return
    if not is_idle(tid):
        # 运行中禁止改
        assign_vars[tid].set(ASSIGN.get(tid, ""))
        log_target("Assignment blocked: target not Idle")
        return
    ASSIGN[tid] = new_wn
    log_target(new_wn, tid, "assignment updated")
    save_config_debounced()
    # 更新一下 resolved path 预览日志（可选）
    path = resolved_shortcut_for(tid)
    log_target(new_wn, tid, f"resolved shortcut → {path or '<unavailable>'}")


# ---- HTTP helper with clearer errors ----

def api(worker_name, path, method="GET", payload=None, timeout=10):
    base = WORKERS[worker_name]
    url = f"{base['url'].rstrip('/')}{path}"
    try:
        if method == "GET":
            r = requests.get(url, timeout=timeout)
        else:
            r = requests.post(url, json=payload or {}, timeout=timeout)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.ConnectTimeout:
        return {"ok": False, "error": f"ConnectTimeout: {url}"}
    except requests.exceptions.ReadTimeout:
        return {"ok": False, "error": f"ReadTimeout: {url}"}
    except requests.exceptions.ConnectionError as e:
        return {"ok": False, "error": f"ConnectionError: {url} ({e.__class__.__name__})"}
    except requests.exceptions.HTTPError as e:
        return {"ok": False, "error": f"HTTPError: {url} ({e})"}
    except Exception as e:
        return {"ok": False, "error": str(e)}

# ---- logging helpers ----

def _human_id(tid: str) -> str:
    try:
        nm = TARGETS.get(str(tid), {}).get("name") or ""
    except Exception:
        nm = ""
    return f"{tid}:{nm}" if nm else str(tid)

def _do_insert_log(line: str, tag: str | None):
    text_log.configure(state="normal")
    text_log.insert(tk.END, line, tag if tag else None)
    text_log.see(tk.END)
    text_log.configure(state="disabled")

def log_target(worker_or_msg=None, target_id=None, text=None):
    """
    通用日志入口（线程安全）：
      - log_target("config saved")
      - log_target("Worker-MSI-Desktop", None, "worker online")
      - log_target("Worker-MSI-Desktop", 3, "launch OK")
    """
    # 兼容仅消息用法：log_target("msg")
    if text is None and target_id is None and isinstance(worker_or_msg, str):
        worker_name = None
        msg = worker_or_msg
        tag = None
        prefix = ""
    else:
        worker_name = worker_or_msg
        # 前缀 & tag
        prefix = ""
        tag = None
        if worker_name is not None:
            prefix += f"[{worker_name}]"
        if target_id is not None:
            tid = str(target_id)
            prefix += f"[{_human_id(tid)}] "
            tag = f"T{tid}"
        elif prefix:
            prefix += " "
        msg = text if text is not None else ""

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    line = f"[{timestamp}] {prefix}{msg}\n"

    if threading.current_thread() is threading.main_thread():
        _do_insert_log(line, tag)
    else:
        root.after(0, lambda: _do_insert_log(line, tag))

def clear_log():
    text_log.configure(state="normal")
    text_log.delete(1.0, tk.END)
    text_log.configure(state="disabled")

# ---- actions ----

def list_all():
    def worker_thread():
        for worker_name in WORKERS.keys():
            res = api(worker_name, "/list", timeout=2)  # list 限制 2s，无重试
            if not res or res.get("ok") is False:
                log_target(worker_name, None, f"/list FAILED → {res}")
            else:
                n = len(res.get("targets", []))
                log_target(worker_name, None, f"/list OK → targets={n}")
    threading.Thread(target=worker_thread, daemon=True).start()


def _log_launch_details(worker_name, tid, res):
    resolved = res.get("resolved") or {}
    if resolved:
        exe = resolved.get("exe"); args = resolved.get("args") or []; cwd = resolved.get("cwd")
        if exe: log_target(worker_name, tid, f"exe: {exe}")
        if args: log_target(worker_name, tid, f"args: {' '.join(args)}")
        if cwd: log_target(worker_name, tid, f"cwd: {cwd}")
    if "window_wait_ms" in res:
        log_target(worker_name, tid, f"wait_window: {res['window_wait_ms']} ms")
    layout = res.get("layout") or {}
    if layout.get("ratio"):
        log_target(worker_name, tid, f"layout(ratio): {layout['ratio']}")
    absr = (layout.get("abs") or {})
    if absr:
        log_target(worker_name, tid, f"layout(abs): x={absr.get('x')}, y={absr.get('y')}, w={absr.get('w')}, h={absr.get('h')}")
    if res.get("handle_close"):
        hc = res["handle_close"]
        log_target(worker_name, tid, f"handle64: ok={hc.get('ok')} msg={hc.get('msg')}")
    for s in (res.get("steps") or []):
        log_target(worker_name, tid, f"· {s}")
    post = res.get("post_launch") or {}
    if post:
        log_target(worker_name, tid, f"post_launch: {post}")

def orchestrate(
    selected_ids: Iterable[str],
    handler: Callable[[str, str], Dict[str, Any]],
    *,
    op_name: str,
    max_parallel_workers: Optional[int] = None,
) -> threading.Thread:
    """
    同一 worker 内按 selected_ids 出现顺序串行，不同 worker 之间并行。
    handler: (worker_name, tid) -> result(dict-like, 期望包含 ok 字段)
    返回最外层 orchestrator 线程对象（daemon）。
    """

    def run_for_worker(worker_name: str, tids_for_worker: List[str]):
        delay = float(WORKERS.get(worker_name, {}).get("join_delay_sec", 0))
        for idx, tid in enumerate(tids_for_worker):
            log_target(worker_name, tid, f"{op_name} start")
            try:
                res = handler(worker_name, tid) or {}
            except Exception as e:
                res = {"ok": False, "error": f"{op_name} error: {e}"}

            if not res or res.get("ok") is False:
                log_target(worker_name, tid, f"{op_name} FAIL → {res}")
            else:
                log_target(worker_name, tid, f"{op_name} OK → " + ", ".join(
                    f"{k}={v}" for k, v in res.items() if k in ("pid", "hwnd", "extra")
                ))

            # 仅在同一 worker 的队列内部，任务之间 sleep
            if delay and (idx + 1) < len(tids_for_worker):
                time.sleep(delay)


    def orchestrator_thread():
        # 1) 按出现顺序分组
        worker_queues: Dict[str, List[str]] = defaultdict(list)
        worker_order: List[str] = []
        for tid in selected_ids:
            wn = ASSIGN.get(tid)
            if not wn:
                log_target("<unknown>", tid, f"{op_name} skipped: no assignment")
                continue
            if wn not in worker_queues:
                worker_order.append(wn)
            worker_queues[wn].append(tid)

        # 2) 控制并发的 worker 数量（可选）
        limit = max_parallel_workers or len(worker_order)
        sem = threading.Semaphore(limit)

        # 3) 为每个 worker 开线程并行，线程内保持顺序
        for wn in worker_order:
            sem.acquire()
            def _start_worker(wn=wn):
                try:
                    run_for_worker(wn, worker_queues[wn])
                finally:
                    sem.release()
            threading.Thread(target=_start_worker, daemon=True).start()

    t = threading.Thread(target=orchestrator_thread, daemon=True)
    t.start()
    return t

def _launch_handler(worker_name: str, tid: str):
    # 组装 payload：优先带上 resolved shortcut_path
    payload = {"target_id": str(tid)}
    try:
        path = resolved_shortcut_for(str(tid))
    except Exception:
        path = None
    if path:
        payload["shortcut_path"] = path
        log_target(worker_name, tid, f"launch with shortcut_path → {path}")
    else:
        log_target(worker_name, tid, "launch without shortcut_path (fallback)")

    # 发起调用（其余逻辑不变）
    return api(worker_name, "/launch", method="POST", payload=payload, timeout=90) or {}


def run_launch(selected_ids):
    def _wrapped_handler(wn, tid):
        tid = str(tid)
        STATE[tid] = "Launching"
        refresh_row_enabled(tid)
        # （可选）打印解析出的完整路径，便于核对
        path = resolved_shortcut_for(tid)
        log_target(wn, tid, f"resolved shortcut → {path or '<unavailable>'}")
        try:
            res = _launch_handler(wn, tid) or {}
        except Exception as e:
            res = {"ok": False, "error": f"{e}"}
        STATE[tid] = "Running" if res.get("ok") else "Error"
        refresh_row_enabled(tid)
        return res

    return orchestrate(
        selected_ids,
        handler=_wrapped_handler,
        op_name="launch",
        max_parallel_workers=None
    )


def _stop_handler(worker_name: str, tid: str) -> Dict[str, Any]:
    res = api(worker_name, "/stop", method="POST", payload={"target_id": tid}, timeout=90) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_stop(selected_ids):
    def _wrapped_handler(wn, tid):
        tid = str(tid)
        STATE[tid] = "Stopping"
        refresh_row_enabled(tid)
        try:
            res = _stop_handler(wn, tid) or {}
        except Exception as e:
            res = {"ok": False, "error": f"{e}"}
        STATE[tid] = "Idle" if res.get("ok") else "Error"
        refresh_row_enabled(tid)
        return res

    return orchestrate(
        selected_ids,
        handler=_wrapped_handler,
        op_name="stop",
        max_parallel_workers=None
    )


# ---- BO handlers ----
def _bo_handler(worker_name: str, tid: str) -> Dict[str, Any]:
    return api(worker_name, "/bo", method="POST", payload={"target_id": tid}, timeout=30) or {}

def run_bo(selected_ids):
    def _wrapped_handler(wn, tid):
        tid = str(tid)
        # bo 是瞬时动作，也把状态临时标记为 Running 以锁编辑
        prev = STATE.get(tid, "Idle")
        STATE[tid] = "Running"
        refresh_row_enabled(tid)
        try:
            res = _bo_handler(wn, tid) or {}
        except Exception as e:
            res = {"ok": False, "error": f"{e}"}
        # bo 完成后：若之前是 Idle 就回 Idle；否则保持 Running（按你需要可调整）
        STATE[tid] = prev if res.get("ok") else "Error"
        refresh_row_enabled(tid)
        return res

    return orchestrate(
        selected_ids,
        handler=_wrapped_handler,
        op_name="bo_command",
        max_parallel_workers=None
    )



def _join_handler(worker_name: str, tid: str, game: str, pwd: str) -> Dict[str, Any]:
    res = api(worker_name, "/join_game", method="POST", payload={"target_id": tid, "game_name": game, "password": pwd}, timeout=60) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_join(selected_ids, game, pwd):
    def _wrapped_handler(wn, tid):
        tid = str(tid)
        # join 期间也算 Running（防止编辑）
        prev = STATE.get(tid, "Idle")
        STATE[tid] = "Running"
        refresh_row_enabled(tid)
        try:
            res = _join_handler(wn, tid, game, pwd) or {}
        except Exception as e:
            res = {"ok": False, "error": f"{e}"}
        # join 失败 → Error；成功保持 Running
        STATE[tid] = "Running" if res.get("ok") else "Error"
        refresh_row_enabled(tid)
        return res

    return orchestrate(
        selected_ids,
        handler=_wrapped_handler,
        op_name="join_game",
        max_parallel_workers=None
    )


def _leave_handler(worker_name: str, tid: str) -> Dict[str, Any]:
    res = api(worker_name, "/leave_game", method="POST", payload={"target_id": tid}, timeout=30) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_leave(selected_ids):
    def _wrapped_handler(wn, tid):
        tid = str(tid)
        STATE[tid] = "Stopping"
        refresh_row_enabled(tid)
        try:
            res = _leave_handler(wn, tid) or {}
        except Exception as e:
            res = {"ok": False, "error": f"{e}"}
        # leave 成功 → 回 Idle；失败 → Error
        STATE[tid] = "Idle" if res.get("ok") else "Error"
        refresh_row_enabled(tid)
        return res

    return orchestrate(
        selected_ids,
        handler=_wrapped_handler,
        op_name="leave_game",
        max_parallel_workers=None
    )


# -------------- UI ----------------
root = tk.Tk()
root.title("D2R Orchestrator")
# 美化 ttk 外观
style = ttk.Style()
try:
    style.theme_use("clam")
except Exception:
    pass
style.configure("TButton", padding=6)
style.configure("Nice.TCombobox", padding=4)
style.map("TButton", relief=[("pressed", "sunken"), ("!pressed", "raised")])

frame_targets = tk.LabelFrame(root, text="Targets")
frame_targets.pack(padx=10, pady=6, fill="x")
frame_targets.grid_rowconfigure(0, minsize=34)

var_checks = {}
def _select_all():
    for v in var_checks.values():
        v.set(True)

def _deselect_all():
    for v in var_checks.values():
        v.set(False)

for i in range(1, 9):
    v = tk.BooleanVar(value=False)
    cb = tk.Checkbutton(frame_targets, text=str(i), variable=v)
    cb.grid(row=0, column=i-1, padx=4, pady=4)
    var_checks[str(i)] = v

row_targets = 0
last_col = len(var_checks) - 1

tk.Button(frame_targets, text="Select All", command=_select_all)\
  .grid(row=row_targets, column=last_col+1, padx=(12,4), pady=0, sticky="w")
tk.Button(frame_targets, text="Deselect All", command=_deselect_all)\
  .grid(row=row_targets, column=last_col+2, padx=4, pady=0, sticky="w")

def build_assignment_panel(parent: tk.Widget):
    frame = tk.LabelFrame(parent, text="Assignment • target → worker")
    frame.pack(padx=10, pady=6, fill="x")

    # 表头
    tk.Label(frame, text="ID", width=4).grid(row=0, column=0, padx=(6,4), pady=4, sticky="w")
    tk.Label(frame, text="Name", width=18, anchor="w").grid(row=0, column=1, padx=(0,10), pady=4, sticky="w")
    tk.Label(frame, text="LNK (filename)", width=34, anchor="w").grid(row=0, column=2, padx=(0,10), pady=4, sticky="w")
    tk.Label(frame, text="Worker", width=24, anchor="w").grid(row=0, column=3, padx=(0,10), pady=4, sticky="w")
    tk.Label(frame, text="Status", width=10, anchor="w").grid(row=0, column=4, padx=(0,10), pady=4, sticky="w")

    # 行
    row = 1
    # 用 TARGETS 的 key 顺序（如需 1..8 固定顺序，按 sorted(int) ）
    for tid in sorted(TARGETS.keys(), key=lambda x: int(x) if str(x).isdigit() else x):
        tinfo = TARGETS.get(tid, {})
        name = tinfo.get("name", "")
        lnk  = tinfo.get("lnk", "")
        wn   = ASSIGN.get(tid, "")

        # ID
        tk.Label(frame, text=str(tid), width=4, anchor="w").grid(row=row, column=0, padx=(6,4), pady=2, sticky="w")
        # Name（只读）
        lab_name = tk.Label(frame, text=name, width=18, anchor="w")
        lab_name.grid(row=row, column=1, padx=(0,10), pady=2, sticky="w")

        # LNK（Idle 才能改）
        lab_lnk = tk.Label(frame, text=lnk, width=34, anchor="w")
        lab_lnk.grid(row=row, column=2, padx=(0,10), pady=2, sticky="w")
        lnk_labels[tid] = lab_lnk

        # Worker（Idle 才能改）
        var_assign = tk.StringVar(value=wn)
        cmb_worker = ttk.Combobox(frame, textvariable=var_assign, values=list(WORKERS.keys()),
                                  width=24, state="readonly")
        cmb_worker.grid(row=row, column=3, padx=(0,10), pady=2, sticky="w")
        var_assign.trace_add("write", lambda *_a, t=tid: _on_assign_change(t))
        assign_vars[tid] = var_assign
        worker_cmbs[tid] = cmb_worker

        # Status
        lab_status = tk.Label(frame, text=STATE.get(tid, "Idle"), width=10, anchor="w")
        lab_status.grid(row=row, column=4, padx=(0,10), pady=2, sticky="w")
        status_labels[tid] = lab_status

        refresh_row_enabled(tid)
        row += 1

    return frame


# === Unified Toolbar: 左侧 Launch/Stop，右侧 下拉 + BO! ===
# 用带边框的 LabelFrame 包裹中间这行
frame_ops_top = tk.LabelFrame(root, text="Controls")   # ← 原来是 tk.Frame(root)
frame_ops_top.pack(padx=10, pady=6, fill="x", ipady=6)

# 左侧：Launch / Stop（保持 tk.Button 样式）
toolbar_left = tk.Frame(frame_ops_top)
toolbar_left.pack(side="left", padx=(0, 8), anchor="w")

tk.Button(
    toolbar_left, text="Launch Selected",
    command=lambda: run_launch([tid for tid, v in var_checks.items() if v.get()])
).pack(side="left", padx=4)

tk.Button(
    toolbar_left, text="Stop Selected",
    command=lambda: run_stop([tid for tid, v in var_checks.items() if v.get()])
).pack(side="left", padx=4)

def _default_bo_target() -> str:
    # 优先：所有标记 GoToRoFReadyForBO=true 的 target（按数字 ID 升序取第一个）
    true_ids = [
        tid for tid, tcfg in (TARGETS or {}).items()
        if isinstance(tcfg, dict) and (tcfg.get("GoToRoFReadyForBO") is True)
    ]
    if true_ids:
        try:
            return str(sorted(true_ids, key=lambda x: int(x) if str(x).isdigit() else x)[0])
        except Exception:
            return str(sorted(true_ids)[0])

    # 次优先：用户偏好里指定（可选）
    pref_tid = (PREFS or {}).get("bo_default_target")
    if pref_tid:
        return str(pref_tid)

    # 回退：原来的固定 "2"
    return "1"

# 右侧容器：按钮样式的“下拉” + BO!
toolbar_right = tk.Frame(frame_ops_top)
toolbar_right.pack(side="right", padx=(0, 8), anchor="e")

bo_target_var = tk.StringVar(value=_default_bo_target())

# 1) 目标选择按钮（按钮外观）
target_btn = tk.Button(toolbar_right, textvariable=bo_target_var, width=3, relief="raised")
target_btn.pack(side="left", padx=(0, 6))

# 2) 弹出菜单（1~8）
target_menu = tk.Menu(root, tearoff=0)
def _set_target(v: str):
    bo_target_var.set(v)
for i in range(1, 9):
    target_menu.add_command(label=str(i), command=lambda v=str(i): _set_target(v))

def _show_target_menu(evt=None):
    # 把菜单贴着按钮下边缘弹出
    x = target_btn.winfo_rootx()
    y = target_btn.winfo_rooty() + target_btn.winfo_height()
    target_menu.tk_popup(x, y)

# 点击按钮弹出菜单
target_btn.configure(command=_show_target_menu)

# 3) BO 按钮（保持 tk.Button 外观）
def on_bo():
    target_id = bo_target_var.get()
    if not target_id:
        messagebox.showwarning("Warn", "请选择一个 target")
        return
    run_bo([target_id])

tk.Button(toolbar_right, text="BO !", command=on_bo).pack(side="left")



frame_join = tk.LabelFrame(root, text="Join / Leave")
frame_join.pack(padx=10, pady=6, fill="x")
frame_join.grid_rowconfigure(0, minsize=34)

tk.Label(frame_join, text="Game Name").grid(row=0, column=0, sticky="w")
entry_game = tk.Entry(frame_join, width=28)
entry_game.grid(row=0, column=1, padx=6)

tk.Label(frame_join, text="Password").grid(row=0, column=2, sticky="w")
entry_pwd = tk.Entry(frame_join, width=20)
entry_pwd.grid(row=0, column=3, padx=6)

def on_join():
    ids = [tid for tid, v in var_checks.items() if v.get()]
    game = entry_game.get().strip()
    pwd = entry_pwd.get().strip()
    if not ids or not game:
        messagebox.showwarning("Warn", "Select targets and input game name.")
        return
    run_join(ids, game, pwd)

btn_join = tk.Button(frame_join, text="Join Selected", command=on_join)
btn_join.grid(row=0, column=4, padx=6)

# 新增：Leave Game 按钮
btn_leave = tk.Button(frame_join, text="Leave Game", command=lambda: run_leave([tid for tid,v in var_checks.items() if v.get()]))
btn_leave.grid(row=0, column=5, padx=6)

frame_ops = tk.Frame(root)
frame_ops.pack(padx=10, pady=6, fill="x")

btn_list = tk.Button(frame_ops, text="List D2R.exe", command=list_all)
btn_list.pack(side="left")

btn_clear = tk.Button(frame_ops, text="Clear Log", command=clear_log)
btn_clear.pack(side="left", padx=(6,0))

frame_log = tk.LabelFrame(root, text="Log")
frame_log.pack(padx=10, pady=6, fill="both", expand=True)
text_log = scrolledtext.ScrolledText(frame_log, height=18, state="disabled")
LOG_FONT = tkfont.nametofont("TkFixedFont").copy()
if "Consolas" in tkfont.families():
    LOG_FONT.configure(family="Consolas")
LOG_FONT.configure(size=10)
text_log.configure(font=LOG_FONT)
for k, color in COLOR_MAP.items():
    text_log.tag_configure(f"T{k}", foreground=color)
text_log.pack(fill="both", expand=True)
frame_assign = build_assignment_panel(root)

def poll_worker_logs():
    def worker_thread():
        while True:
            for worker_name in WORKERS.keys():
                res = api(worker_name, "/drain_logs", method="POST", payload={"max_items": 200}, timeout=2)
                if not res or res.get("events") is None:
                    continue
                for ev in res["events"]:
                    tid = ev.get("target_id", "?")
                    msg = ev.get("msg", "")
                    log_target(worker_name, tid, msg)
            time.sleep(0.3)
    threading.Thread(target=worker_thread, daemon=True).start()

# 在创建完 text_log、配置好颜色 tag 之后调用：
poll_worker_logs()

root.mainloop()