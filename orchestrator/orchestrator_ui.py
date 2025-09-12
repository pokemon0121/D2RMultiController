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
import requests
from typing import Callable, Dict, Any, Iterable, List, Optional
from functools import partial

CFG_PATH = Path(__file__).with_name("config.json")
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

def load_config():
    if not CFG_PATH.exists():
        messagebox.showerror("Error", f"Missing config: {CFG_PATH}")
        return None
    return json.loads(CFG_PATH.read_text(encoding="utf-8"))

cfg = load_config()
if cfg is None:
    raise SystemExit(1)

WORKERS = cfg["workers"]
TARGET_ASSIGN = cfg["target_assign"]
JOIN_DELAY = float(cfg.get("join_delay_sec", 9))
JOIN_JITTER = float(cfg.get("join_jitter_sec", 3))

# ---- HTTP helper with clearer errors ----

def api(worker_name, path, method="GET", payload=None, timeout=10):
    base = WORKERS[worker_name]
    url = f"{base}{path}"
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

def log(msg, tag=None):
    text_log.configure(state="normal")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    if tag:
        text_log.insert(tk.END, f"[{timestamp}] {msg}\n", tag)
    else:
        text_log.insert(tk.END, f"[{timestamp}] {msg}\n")
    text_log.see(tk.END)
    text_log.configure(state="disabled")

def log_target(worker_name, target_id, text):
    tid = str(target_id)
    prefix = f"[{worker_name}][target {tid}] "
    tag = f"T{tid}"
    log(prefix + text, tag=tag)


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
                log(f"[{worker_name}] /list FAILED → {res}")
            else:
                n = len(res.get("targets", []))
                log(f"[{worker_name}] /list OK → targets={n}")
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
    per_worker_delay: float = 10.0,
    max_parallel_workers: Optional[int] = None,
) -> threading.Thread:
    """
    同一 worker 内按 selected_ids 出现顺序串行，不同 worker 之间并行。
    handler: (worker_name, tid) -> result(dict-like, 期望包含 ok 字段)
    返回最外层 orchestrator 线程对象（daemon）。
    """

    def run_for_worker(worker_name: str, tids_for_worker: List[str]):
        for tid in tids_for_worker:
            log_target(worker_name, tid, f"{op_name} start")
            try:
                res = handler(worker_name, tid) or {}
            except Exception as e:
                res = {"ok": False, "error": f"{op_name} error: {e}"}

            if not res or res.get("ok") is False:
                log_target(worker_name, tid, f"{op_name} FAIL → {res}")
            else:
                log_target(
                    worker_name,
                    tid,
                    f"{op_name} OK → " + ", ".join(
                        f"{k}={v}" for k, v in res.items() if k in ("pid", "hwnd", "extra")
                    )
                )
            time.sleep(per_worker_delay)

    def orchestrator_thread():
        # 1) 按出现顺序分组
        worker_queues: Dict[str, List[str]] = defaultdict(list)
        worker_order: List[str] = []
        for tid in selected_ids:
            wn = TARGET_ASSIGN.get(tid)
            if not wn:
                log_target("<unknown>", tid, f"{op_name} skipped: no TARGET_ASSIGN entry")
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

def _launch_handler(worker_name: str, tid: str) -> Dict[str, Any]:
    res = api(worker_name, "/launch", method="POST", payload={"target_id": tid}, timeout=90) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_launch(selected_ids):
    return orchestrate(
        selected_ids,
        handler=_launch_handler,
        op_name="launch",
        per_worker_delay=10.0,
        max_parallel_workers=None  # 或者 2 / 3 做节流
    )

def _stop_handler(worker_name: str, tid: str) -> Dict[str, Any]:
    res = api(worker_name, "/stop", method="POST", payload={"target_id": tid}, timeout=90) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_stop(selected_ids):
    return orchestrate(
        selected_ids,
        handler=_stop_handler,
        op_name="stop",
        per_worker_delay=0.2,
        max_parallel_workers=None  # 或者 2 / 3 做节流
    )

def _join_handler(worker_name: str, tid: str, game: str, pwd: str) -> Dict[str, Any]:
    res = api(worker_name, "/join_game", method="POST", payload={"target_id": tid, "game_name": game, "password": pwd}, timeout=60) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_join(selected_ids, game, pwd):
    return orchestrate(
        selected_ids,
        handler=partial(_join_handler, game=game, pwd=pwd),
        op_name="join_game",
        per_worker_delay=JOIN_DELAY,
        max_parallel_workers=None  # 或者 2 / 3 做节流
    )

def _leave_handler(worker_name: str, tid: str) -> Dict[str, Any]:
    res = api(worker_name, "/leave_game", method="POST", payload={"target_id": tid}, timeout=30) or {}
    if res.get("ok"):
        _log_launch_details(worker_name, tid, res)
    return res

def run_leave(selected_ids):
    return orchestrate(
        selected_ids,
        handler=_leave_handler,
        op_name="leave_game",
        per_worker_delay=0.5,
        max_parallel_workers=None  # 或者 2 / 3 做节流
    )

# -------------- UI ----------------
root = tk.Tk()
root.title("D2R Orchestrator")

frame_targets = tk.LabelFrame(root, text="Targets")
frame_targets.pack(padx=10, pady=6, fill="x")

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

frame_ops_top = tk.Frame(root)
frame_ops_top.pack(padx=10, pady=6, fill="x")
tk.Button(frame_ops_top, text="Launch Selected", command=lambda: run_launch([tid for tid,v in var_checks.items() if v.get()])).pack(side="left", padx=4)
tk.Button(frame_ops_top, text="Stop Selected", command=lambda: run_stop([tid for tid,v in var_checks.items() if v.get()])).pack(side="left", padx=4)

frame_join = tk.LabelFrame(root, text="Join / Leave")
frame_join.pack(padx=10, pady=6, fill="x")

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