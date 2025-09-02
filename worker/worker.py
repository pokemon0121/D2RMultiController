import argparse
import json
import time
import subprocess
import threading
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set

import psutil
from fastapi import FastAPI
from fastapi import Body
from pydantic import BaseModel
import uvicorn

from collections import deque
from threading import Lock

# Win32
import win32gui
import win32con
import win32api
import win32process
import win32com.client  # resolve .lnk
import ctypes
from ctypes import wintypes
import pythoncom

user32 = ctypes.windll.user32

# 1) 设为 Per-Monitor-V2
ok = ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2
# 2) 验证
aw = ctypes.windll.user32.GetAwarenessFromDpiAwarenessContext(
    ctypes.windll.user32.GetThreadDpiAwarenessContext()
)
# aw==3 才是 Per-Monitor（V1/V2），否则就再用 8.1 的 API 兜底
if not ok or aw != 3:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        # 最差也用 system-aware（不完美，但统一成一套）
        ctypes.windll.user32.SetProcessDPIAware()
        
# ============== Admin / UAC helpers ==============

def _is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

# ============== Global event queue for UI polling ==============
LOGQ = deque(maxlen=3000)
LOGQ_LOCK = Lock()

def log_event(target_id: str | None, msg: str):
    with LOGQ_LOCK:
        LOGQ.append({"target_id": str(target_id), "msg": msg})

def sleep_log(target_id: Optional[str], seconds: float):
    if seconds > 1.0 and target_id is not None:
        log_event(target_id, f"sleep {seconds:.1f}s")
    time.sleep(seconds)

# ---- Unified return helper (always logs) ----

def log_and_return(target_id: Optional[str], payload: dict):
    try:
        log_event(target_id, f"return: {payload}")
    except Exception:
        pass
    return payload

# ============== Helpers: Windows / Input ==============

# ============== Window Rect Utils ==============

_AdjustWindowRectEx = user32.AdjustWindowRectEx
_AdjustWindowRectEx.argtypes = [
    wintypes.LPRECT,      # LPRECT
    wintypes.DWORD,       # style
    wintypes.BOOL,        # bMenu
    wintypes.DWORD,       # exstyle
]
_AdjustWindowRectEx.restype = wintypes.BOOL

class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]

def get_screen_wh() -> tuple[int, int]:
    try:
        return win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
    except Exception:
        return (1920, 1080)

def apply_window_rect_ratio_abs(hwnd: int, cfg: dict, target_id: Optional[str] = None):
    """根据比例配置计算绝对像素并移动窗口，返回 (x,y,w,h)。"""
    if not cfg:
        return None
    sw, sh = get_screen_wh()
    try:
        xr = float(cfg.get("x", 0.0))
        yr = float(cfg.get("y", 0.0))
        wr = float(cfg.get("w", 1.0))
        hr = float(cfg.get("h", 1.0))
    except Exception:
        return None
    x = int(round(xr * sw))
    y = int(round(yr * sh))
    w = max(1, int(round(wr * sw)))
    h = max(1, int(round(hr * sh)))
    try:
        if target_id is not None:
            log_event(target_id, f"move window → x={x} y={y} w={w} h={h}")
        win32gui.MoveWindow(hwnd, x, y, w, h, True)
    except Exception:
        pass
    return {"x": x, "y": y, "w": w, "h": h}

def is_minimized(hwnd: int) -> bool:
    return win32gui.IsIconic(hwnd) != 0

def ensure_restored_no_focus(hwnd: int):
    if is_minimized(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

def get_window_rect(hwnd: int) -> Dict[str, int]:
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    return {"left": l, "top": t, "right": r, "bottom": b}

# ---- DPI helpers: enforce physical pixels end-to-end ----

def dpi_scale(hwnd: int) -> float:
    try:
        dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
        return float(dpi) / 96.0 if dpi else 1.0
    except Exception:
        return 1.0

def get_client_size(hwnd: int) -> Tuple[int, int]:
    l, t, r, b = win32gui.GetClientRect(hwnd)
    return (r - l, b - t)


# client(physical) -> screen(physical)

def client_phys_to_screen_phys(hwnd: int, x_phys: int, y_phys: int) -> Tuple[int, int]:
    return win32gui.ClientToScreen(hwnd, (int(x_phys), int(y_phys)))



# 1) 修改打包 lParam：直接打包“client 物理像素”，不要再除 DPI
def make_lparam_client_phys(hwnd: int, x_client_phys: int, y_client_phys: int) -> int:
    return ((int(y_client_phys) & 0xFFFF) << 16) | (int(x_client_phys) & 0xFFFF)

def get_client_origin_abs_phys(hwnd: int) -> Tuple[int,int]:
    return client_phys_to_screen_phys(hwnd, 0, 0)

def log_client_geom(hwnd: int, target_id: str | None = None):
    w, h = get_client_size(hwnd)
    ox, oy = get_client_origin_abs_phys(hwnd)
    log_event(target_id, f"client: size={w}x{h} origin@{ox},{oy}")

# ============== UiMap point → abs(screen) ==============

def get_physical_desktop_wh() -> Tuple[int, int]:
    dm = win32api.EnumDisplaySettings(None, win32con.ENUM_CURRENT_SETTINGS)
    return dm.PelsWidth, dm.PelsHeight

# 2)（可选，仅用于对齐日志）在点击时同时打印 client 的逻辑坐标，便于和 C# 样式比对
def _hwnd_chain_contains(parent: int, child: int) -> bool:
    try:
        h = child
        while h and win32gui.IsWindow(h):
            if int(h) == int(parent):
                return True
            h = win32gui.GetParent(h)
    except Exception:
        pass
    return False

def uimap_point_client(hwnd: int, spec) -> Tuple[int,int]:
    """支持 [rx, ry] 百分比或 (x, y) 像素（client-physical），返回 client-physical。"""
    w, h = get_client_size(hwnd)
    try:
        x, y = spec
    except Exception:
        raise ValueError("UiMap point must be a 2-tuple/list of numbers")
    if isinstance(x, (int, float)) and isinstance(y, (int, float)) and 0.0 <= float(x) <= 1.0 and 0.0 <= float(y) <= 1.0:
        cx = int(float(x) * max(w - 1, 1))
        cy = int(float(y) * max(h - 1, 1))
    else:
        cx = int(round(float(x)))
        cy = int(round(float(y)))
    return cx, cy

def ensure_hwnd(target_id: str, cur_hwnd: int | None = None) -> int:
    """点击前确保拿到可用的 hwnd；若当前无效则 refresh_targets 重取。"""
    try:
        if cur_hwnd and win32gui.IsWindow(cur_hwnd):
            return cur_hwnd
    except Exception:
        pass
    # 用你的目标注册表刷新
    refresh_targets()
    hwnd = TARGET_MAP.get(target_id)
    if not hwnd or not win32gui.IsWindow(hwnd):
        raise RuntimeError(f"target {target_id} not found or invalid hwnd")
    return hwnd

def bg_mouse_click_client(hwnd: int, cx: int, cy: int, target_id: str | None = None):
    # 0) 点击前校验/恢复 hwnd
    if target_id:
        try:
            prev = hwnd
        except Exception:
            prev = 0
        hwnd = ensure_hwnd(target_id, cur_hwnd=hwnd)
        if prev and prev != hwnd:
            log_event(target_id, f"hwnd refresh: {prev:08X} -> {hwnd:08X}")
    else:
        if not win32gui.IsWindow(hwnd):
            raise RuntimeError("invalid hwnd and no target_id to recover")

    # 1) client -> screen（这一步就不会再抛 1400 了）
    absx, absy = client_phys_to_screen_phys(hwnd, cx, cy)

    # 2) 只移动鼠标 + 只发 WinMsg（不前台）
    win32api.SetCursorPos((absx, absy))
    # 命中窗口优先（同 root 时）
    hit = win32gui.WindowFromPoint((absx, absy))
    GA_ROOT = 2
    def _root(h): 
        try: return win32gui.GetAncestor(h, GA_ROOT)
        except Exception: return 0
    th, tx, ty = hwnd, cx, cy
    if hit and hit != hwnd and _root(hit) == _root(hwnd):
        tx, ty = win32gui.ScreenToClient(hit, (absx, absy))
        th = hit

    lp = (int(ty) << 16) | (int(tx) & 0xFFFF)
    win32gui.SendMessage(th, win32con.WM_MOUSEMOVE,   0, lp)
    win32gui.SendMessage(th, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
    win32gui.SendMessage(th, win32con.WM_LBUTTONUP,   0, lp)

    if target_id:
        log_event(target_id, f"winmsg+cursor: sent to {th:08X} at client@{tx},{ty}; abs@{absx},{absy}")


def bg_send_char(hwnd: int, ch: str, target_id: Optional[str] = None):
    if target_id is not None:
        log_event(target_id, f"char '{ch}'")
    win32api.PostMessage(hwnd, win32con.WM_CHAR, ord(ch), 0)

def bg_send_hotkey(hwnd: int, vks: List[int], hold_ms: int = 30, target_id: Optional[str] = None):
    """
    发送组合键，例如 Ctrl+A:
        bg_send_hotkey(hwnd, [win32con.VK_CONTROL, 0x41])
    单键：
        bg_send_hotkey(hwnd, [win32con.VK_SPACE])
    """
    if target_id is not None:
        log_event(target_id, f"hotkey {vks} down ({hold_ms}ms)")
    for vk in vks:
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
    if hold_ms and hold_ms > 0:
        time.sleep(hold_ms / 1000.0)
    for vk in reversed(vks):
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)
    if target_id is not None:
        log_event(target_id, f"hotkey {vks} up")

# ============== UI helpers ==============

def ui_press_space(hwnd: int, target_id: Optional[str] = None):
    if target_id is not None:
        log_event(target_id, "press SPACE")
    bg_send_hotkey(hwnd, [win32con.VK_SPACE], target_id=target_id)

def ui_press_enter(hwnd: int, target_id: Optional[str] = None):
    if target_id is not None:
        log_event(target_id, "press ENTER")
    bg_send_hotkey(hwnd, [win32con.VK_RETURN], target_id=target_id)

def ui_press_esc(hwnd: int, target_id: Optional[str] = None):
    if target_id is not None:
        log_event(target_id, "press ESC")
    bg_send_hotkey(hwnd, [win32con.VK_ESCAPE], target_id=target_id)

# ============== UiMap loader ==============

def load_uimap(path: str):
    p = Path(path)
    if not p.exists():
        with LOGQ_LOCK:
            LOGQ.append({"target_id": "system", "msg": f"uimap: file not found: {path}"})
        return {}
    try:
        with p.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        with LOGQ_LOCK:
            LOGQ.append({"target_id": "system", "msg": f"uimap: parse error: {e}"})
        return {}
    out = {}
    for k, v in data.items():
        if isinstance(v, (list, tuple)) and len(v) == 2:
            out[k] = (float(v[0]), float(v[1]))
    return out

import win32con, win32gui, win32api

GA_PARENT = 1
GA_ROOT   = 2

def _get_class(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd)
    except Exception:
        return "?"

def _dump_hwnd_chain(hwnd: int) -> str:
    chain = []
    try:
        cur = hwnd
        while cur:
            try:
                title = win32gui.GetWindowText(cur)
            except Exception:
                title = ""
            cls = _get_class(cur)
            chain.append(f"{cur:08X}[{cls}] '{title}'")
            cur = win32gui.GetParent(cur)
    except Exception:
        pass
    return " -> ".join(chain)

def _root(hwnd: int) -> int:
    try:
        return win32gui.GetAncestor(hwnd, GA_ROOT)
    except Exception:
        return 0

def click_with_probe(target_id: str, hwnd: int, cx: int, cy: int):
    # client -> abs
    absx, absy = client_phys_to_screen_phys(hwnd, cx, cy)

    # 命中谁
    hit = win32gui.WindowFromPoint((absx, absy))

    # 先发给目标 hwnd（MOVE -> DOWN -> UP）
    lp = make_lparam_client_phys(hwnd, cx, cy)
    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lp)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)

    # 如果命中的子窗口不同，再发一遍给“命中窗口”
    if hit and hit != hwnd:
        cx2, cy2 = win32gui.ScreenToClient(hit, (absx, absy))
        lp2 = (cy2 << 16) | (cx2 & 0xFFFF)
        win32gui.SendMessage(hit, win32con.WM_MOUSEMOVE, 0, lp2)
        win32gui.SendMessage(hit, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp2)
        win32gui.SendMessage(hit, win32con.WM_LBUTTONUP, 0, lp2)
        log_event(target_id, f"probe: re-sent to child {hit:08X} at client@{cx2},{cy2}")


def debug_probe_full(target_id: str, hwnd: int, cx: int, cy: int):
    # 1) client 基本信息
    w_cli, h_cli = get_client_size(hwnd)                 # GetClientRect
    ox_abs, oy_abs = get_client_origin_abs_phys(hwnd)    # ClientToScreen(0,0)
    log_event(target_id, f"client_wh={w_cli}x{h_cli} origin_abs={ox_abs},{oy_abs} hwnd={hwnd:08X}")

    # 2) client -> screen 两条路径（应一致）
    abs1x, abs1y = client_phys_to_screen_phys(hwnd, cx, cy)
    abs2x, abs2y = ox_abs + cx, oy_abs + cy
    log_event(target_id, f"abs_via_ClientToScreen={abs1x},{abs1y}  abs_via_origin+offset={abs2x},{abs2y}")

    # 3) 用物理分辨率判断是否在屏内
    phys_w, phys_h = get_physical_desktop_wh()
    vx, vy = 0, 0  # 如果你是多屏且有负坐标，再用 SM_XVIRTUALSCREEN/SM_YVIRTUALSCREEN 当原点
    log_event(target_id, f"vdesk(physical) origin@{vx},{vy} size={phys_w}x{phys_h}")
    in_vdesk = (vx <= abs1x < vx + phys_w) and (vy <= abs1y < vy + phys_h)
    log_event(target_id, f"abs_in_vdesk={in_vdesk}")

    # 4) 命中窗口（两条 abs 都测一下）
    hit1 = win32gui.WindowFromPoint((abs1x, abs1y)) if in_vdesk else None
    hit2 = win32gui.WindowFromPoint((abs2x, abs2y)) if (vx <= abs2x < vx + phys_w) and (vy <= abs2y < vy + phys_h) else None
    log_event(target_id, f"hit1={hit1 and f'{hit1:08X}'} hit2={hit2 and f'{hit2:08X}'}")

    # 5) 打印窗口链，便于确认父/子关系
    def dump_chain(h):
        chain = []
        cur = h
        while cur:
            try:
                cls = win32gui.GetClassName(cur)
                ttl = win32gui.GetWindowText(cur)
            except Exception:
                cls, ttl = "?", ""
            chain.append(f"{cur:08X}[{cls}] '{ttl}'")
            cur = win32gui.GetParent(cur)
        return " -> ".join(chain)

    if hit1:
        log_event(target_id, "chain@abs1: " + dump_chain(hit1))
    if hit2 and hit2 != hit1:
        log_event(target_id, "chain@abs2: " + dump_chain(hit2))


# ============== API + Helper: Goto Lobby ==============
class GotoLobbyReq(BaseModel):
    target_id: str

def _do_goto_lobby(target_id: str, hwnd: int) -> dict:
    log_event(target_id, "goto_lobby: start")
    ensure_restored_no_focus(hwnd)
    ui = load_uimap(UIMAP_PATH)

    log_event(target_id, f"goto_lobby: uimap_path={UIMAP_PATH}")
    log_event(target_id, f"goto_lobby: OnlineTab raw={ui.get('OnlineTab')}, GoToLobbyButton raw={ui.get('GoToLobbyButton')}")

    if "OnlineTab" not in ui or "GoToLobbyButton" not in ui:
        log_event(target_id, f"goto_lobby: missing keys; ui_keys={list(ui.keys())}")
        return log_and_return(target_id, {"ok": False, "error": "UiMap missing OnlineTab/GoToLobbyButton"})

    w, h = get_client_size(hwnd)
    log_client_geom(hwnd, target_id)

    rx, ry = ui["OnlineTab"]
    cx = int(rx * max(w - 1, 1)) if 0 <= rx <= 1 and 0 <= ry <= 1 else int(rx)
    cy = int(ry * max(h - 1, 1)) if 0 <= rx <= 1 and 0 <= ry <= 1 else int(ry)
    log_event(target_id, f"goto_lobby: OnlineTab client@{cx},{cy} from ratio@{rx},{ry}")

    cx, cy = uimap_point_client(hwnd, tuple(ui["OnlineTab"]))
    bg_mouse_click_client(hwnd, cx, cy, target_id=target_id)

    sleep_log(target_id, 1.0)

    rx, ry = ui["GoToLobbyButton"]
    cx = int(rx * max(w - 1, 1)) if 0 <= rx <= 1 and 0 <= ry <= 1 else int(rx)
    cy = int(ry * max(h - 1, 1)) if 0 <= rx <= 1 and 0 <= ry <= 1 else int(ry)
    log_event(target_id, f"goto_lobby: GoToLobbyButton client@{cx},{cy} from ratio@{rx},{ry}")

    cx, cy = uimap_point_client(hwnd, tuple(ui["GoToLobbyButton"]))
    bg_mouse_click_client(hwnd, cx, cy, target_id=target_id)

    log_event(target_id, "goto_lobby: done")
    return log_and_return(target_id, {"ok": True, "steps": ["click OnlineTab", "sleep 1s", "click GoToLobbyButton"]})

app = FastAPI()

# ---- Global unhandled exception logger ----
from fastapi.responses import JSONResponse
from fastapi.requests import Request

@app.exception_handler(Exception)
async def _unhandled_exc(request: Request, exc: Exception):
    try:
        body = await request.json()
        target_id = body.get("target_id") if isinstance(body, dict) else None
    except Exception:
        target_id = None
    log_event(target_id or "system", f"unhandled: {type(exc).__name__}: {exc}")
    return JSONResponse(status_code=500, content={"ok": False, "error": str(exc)})

@app.post("/goto_lobby")
def goto_lobby(req: GotoLobbyReq):
    target_id = req.target_id
    hwnd = TARGET_MAP.get(target_id)
    if not hwnd:
        refresh_targets()
        hwnd = TARGET_MAP.get(target_id)
        if not hwnd:
            return log_and_return(target_id, {"ok": False, "error": "target not found"})
    ensure_restored_no_focus(hwnd)
    return _do_goto_lobby(target_id, hwnd)

# ============== Shortcut / Launch helpers ==============

def _resolve_lnk_mcybe(path_str: str):
    p = Path(path_str)
    cand = [p]
    if p.suffix.lower() != ".lnk":
        cand.append(p.with_suffix(p.suffix + ".lnk" if p.suffix else ".lnk"))
    for c in cand:
        if c.exists() and c.is_file() and c.suffix.lower() == ".lnk":
            pythoncom.CoInitialize()
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                sc = shell.CreateShortcut(str(c))
                return sc.Targetpath or "", sc.Arguments or "", sc.WorkingDirectory or ""
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    return None

def _resolve_lnk_maybe(path_str: str):
    p = Path(path_str)
    cand = [p]
    if p.suffix.lower() != ".lnk":
        cand.append(p.with_suffix(p.suffix + ".lnk" if p.suffix else ".lnk"))
    for c in cand:
        if c.exists() and c.is_file() and c.suffix.lower() == ".lnk":
            pythoncom.CoInitialize()
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                sc = shell.CreateShortcut(str(c))
                return sc.Targetpath or "", sc.Arguments or "", sc.WorkingDirectory or ""
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    return None


def resolve_launcher(cfg: dict):
    args_append = (cfg.get("args_append") or "").strip()

    sc = cfg.get("shortcut")
    if sc:
        info = _resolve_lnk_maybe(sc)
        if info:
            target, base_args, cwd = info
            final_args = (base_args + " " + args_append).strip()
            cmd = [target] + ([a for a in final_args.split(" ") if a] if final_args else [])
            return cmd, (cwd or None)

    exep = cfg.get("exe")
    if exep:
        target = str(Path(exep))
        final_args = args_append
        cmd = [target] + ([a for a in final_args.split(" ") if a] if final_args else [])
        return cmd, None

    dr = cfg.get("dir")
    if dr:
        d2r = Path(dr) / "D2R.exe"
        if d2r.exists():
            target = str(d2r)
            final_args = args_append
            cmd = [target] + ([a for a in final_args.split(" ") if a] if final_args else [])
            return cmd, str(Path(dr))

    raise RuntimeError("Invalid launcher config: need shortcut/exe/dir")

def enumerate_windows_by_title(substr: str) -> List[int]:
    result = []
    lower = substr.lower()
    def cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd) or ""
            if lower in title.lower():
                result.append(hwnd)
        except:
            pass
    win32gui.EnumWindows(cb, None)
    return result

def find_top_window_for_pid(pid: int) -> Optional[int]:
    hwnds = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid == pid:
                hwnds.append(hwnd)
    win32gui.EnumWindows(cb, None)
    titled = [h for h in hwnds if win32gui.GetWindowText(h)]
    return titled[0] if titled else (hwnds[0] if hwnds else None)

import time
import win32con
import win32clipboard as cb

def set_clipboard_text(text: str, retries: int = 5, delay_ms: int = 50):
    """把字符串写入系统剪贴板（CF_UNICODETEXT），带重试，避免“剪贴板被占用”报错。"""
    for i in range(retries):
        try:
            cb.OpenClipboard()
            cb.EmptyClipboard()
            cb.SetClipboardData(win32con.CF_UNICODETEXT, text)
            cb.CloseClipboard()
            return True
        except Exception:
            try:
                cb.CloseClipboard()
            except Exception:
                pass
            if i == retries - 1:
                raise
            time.sleep(delay_ms / 1000.0)


# ---- PID resolution ----

def _existing_assigned_pids() -> Set[int]:
    return {pid for pid in TARGET_PID.values() if pid}

def _iter_descendants(proc: psutil.Process):
    try:
        for child in proc.children(recursive=True):
            yield child
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return

def refresh_targets():
    # 清理以及重建
    for tid in list(TARGET_MAP.keys()):
        hwnd = TARGET_MAP.get(tid)
        try:
            if not hwnd or not win32gui.IsWindow(hwnd):
                TARGET_MAP.pop(tid, None)
        except Exception:
            TARGET_MAP.pop(tid, None)

    for tid, pid in list(TARGET_PID.items()):
        try:
            if not pid or not psutil.pid_exists(pid):
                TARGET_PID.pop(tid, None)
                TARGET_MAP.pop(tid, None)
        except Exception:
            TARGET_PID.pop(tid, None)
            TARGET_MAP.pop(tid, None)

    for tid, pid in list(TARGET_PID.items()):
        try:
            hwnd = find_top_window_for_pid(pid)
            if hwnd and win32gui.IsWindow(hwnd):
                TARGET_MAP[tid] = hwnd
        except Exception:
            pass

    try:
        ws = enumerate_windows_by_title(WINDOW_TITLE_MATCH)
    except Exception:
        ws = []
    idx = 1
    for hwnd in ws:
        tid = str(idx)
        if tid not in TARGET_MAP:
            try:
                if win32gui.IsWindow(hwnd):
                    TARGET_MAP[tid] = hwnd
            except Exception:
                pass
        idx += 1

def discover_d2r_pids():
    pids = []
    for p in psutil.process_iter(["pid", "name", "create_time"]):
        try:
            name = (p.info.get("name") or "").lower()
            if name == "d2r.exe":
                pids.append((p.info.get("create_time", 0.0), p.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    pids.sort(key=lambda x: x[0])
    return [pid for _, pid in pids]

def _find_final_pid(proc: Optional[subprocess.Popen], exe_basename: Optional[str], start_time: float) -> Optional[int]:
    if proc and getattr(proc, "pid", None):
        try:
            p = psutil.Process(proc.pid)
            if p.is_running():
                return proc.pid
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if proc and getattr(proc, "pid", None):
        try:
            parent = psutil.Process(proc.pid)
            cand = []
            for ch in _iter_descendants(parent):
                try:
                    if exe_basename and ch.name().lower() != exe_basename.lower():
                        continue
                    if ch.create_time() >= start_time - 1:
                        cand.append((ch.create_time(), ch.pid))
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            cand.sort()
            for _, pid in reversed(cand):
                if pid not in _existing_assigned_pids():
                    return pid
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if exe_basename:
        cand = []
        for p in psutil.process_iter(["pid", "name", "create_time"]):
            try:
                if not p.info["name"] or p.info["name"].lower() != exe_basename.lower():
                    continue
                if p.info["create_time"] < start_time - 1:
                    continue
                if p.info["pid"] in _existing_assigned_pids():
                    continue
                cand.append((p.info["create_time"], p.info["pid"]))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        cand.sort()
        if cand:
            return cand[-1][1]
    return None

# ============== Handle Close (Sysinternals handle64) ==============

def close_other_instance_handle(pid: int) -> (bool, str):
    """
    Kill other instance handle locks using Sysinternals handle64.exe.
    Returns (ok, msg).
    """
    try:
        exe_paths = [
            Path(__file__).parent / "Utility" / "handle64.exe",
            Path("Utility") / "handle64.exe"
        ]
        tool = None
        for p in exe_paths:
            if p.exists():
                tool = str(p)
                break
        if not tool:
            return (False, "ToolNotFound: handle64.exe missing in Utility/")
        # Must run as Administrator
        if not ctypes.windll.shell32.IsUserAnAdmin():
            return (False, "AdminRequired: run Worker as Administrator for handle64.exe")

        # Run handle64
        out = subprocess.check_output([tool, "-p", str(pid), "-v", "-a"], text=True, stderr=subprocess.STDOUT)
        # Close all handles except current
        closed = 0
        for line in out.splitlines():
            if "DiabloII Check For Other Instances" in line:
                parts = line.split(",")
                if len(parts) > 3:
                    handle_hex = parts[3].strip()
                    try:
                        subprocess.run([tool, "-p", str(pid), "-c", handle_hex, "-y"], check=True)
                        closed += 1
                    except Exception as e:
                        return (False, f"CloseFailed: {e}")
        return (True, f"Closed {closed} handles")
    except subprocess.CalledProcessError as e:
        return (False, f"handle64 error: {e.output}")
    except Exception as e:
        return (False, f"Exception: {e}")

# ============== Worker State ==============
TARGET_MAP: Dict[str, int] = {}
TARGET_PID: Dict[str, int] = {}
WORKER_NAME: str = "Worker-A"
WINDOW_TITLE_MATCH: str = "Diablo II"
UIMAP_PATH: str = str(Path(__file__).parent/"uimaps"/"default.json")
LAUNCHERS: Dict[str, dict] = {}
POST_LAUNCH: dict = {"enabled": True, "sequence": "default"}
CLOSE_HANDLE_AFTER_LAUNCH: bool = True
JOIN_LOCK = threading.Lock()

# Aspect lock defaults (可被 config.json 覆盖)
LOCK_ASPECT: bool = False
ASPECT_W: int = 8
ASPECT_H: int = 5
ASPECT_ANCHOR: str = "topleft"   # "topleft" | "center"
ASPECT_MODE: str = "auto"        # "auto" | "width" | "height"

# ============== API Models ==============

class FocusReq(BaseModel):
    target_id: str

class ClickReq(BaseModel):
    target_id: str
    cx: int
    cy: int

class TypeReq(BaseModel):
    target_id: str
    text: str

class JoinReq(BaseModel):
    target_id: str
    game_name: str
    password: Optional[str] = ""

class LeaveReq(BaseModel):
    target_id: str

class LaunchReq(BaseModel):
    target_id: str

class StopReq(BaseModel):
    target_id: str
    force: bool = False

class CloseHandleReq(BaseModel):
    target_id: str

# ============== FastAPI App ==============

app = FastAPI()

@app.get("/name")
def name():
    return {"worker": WORKER_NAME}

@app.get("/admin_status")
def admin_status():
    return {"is_admin": _is_admin()}

@app.get("/pidmap")
def pidmap():
    return {"pidmap": TARGET_PID}

@app.get("/list")
def list_targets():
    try:
        d2r_pids = discover_d2r_pids()
        targets = [{"id": str(i + 1), "pid": pid} for i, pid in enumerate(d2r_pids)]
        return {"worker": WORKER_NAME, "targets": targets}
    except Exception as e:
        return {"worker": WORKER_NAME, "ok": False, "error": f"list_failed: {e}"}

@app.post("/close_handle")
def close_handle(req: CloseHandleReq):
    pid = TARGET_PID.get(req.target_id)
    if not pid:
        return log_and_return(req.target_id, {"ok": False, "error": "no known pid for target"})
    ok, msg = close_other_instance_handle(pid)
    return log_and_return(req.target_id, {"ok": ok, "msg": msg})

@app.post("/launch")
def launch(req: LaunchReq):
    cfg = LAUNCHERS.get(req.target_id, {})
    if not cfg:
        return log_and_return(req.target_id, {"ok": False, "error": f"No launcher configured for target {req.target_id}."})

    debug_steps = []
    t0 = time.time()

    try:
        cmd, cwd = resolve_launcher(cfg)
        debug_steps.append(f"resolve: exe={cmd[0]}")
        if len(cmd) > 1:
            debug_steps.append("resolve: args=" + " ".join(cmd[1:]))
        if cwd:
            debug_steps.append(f"resolve: cwd={cwd}")
        debug_steps.append("subprocess: launching ...")
        proc = subprocess.Popen(cmd, cwd=cwd or None)
        debug_steps.append(f"subprocess: pid={proc.pid}")
    except Exception as e:
        return log_and_return(req.target_id, {"ok": False, "error": f"launch failed: {e}"})

    exe_basename = Path(cmd[0]).name if cmd and cmd[0] else None

    final_pid = _find_final_pid(proc, exe_basename, t0)
    debug_steps.append(f"pid resolution: exe_basename={exe_basename} final_pid={final_pid}")
    TARGET_PID[req.target_id] = final_pid

    hwnd = None
    wait_deadline = time.time() + 20.0
    while time.time() < wait_deadline:
        if final_pid:
            hwnd = find_top_window_for_pid(final_pid)
            if hwnd:
                break
        sleep_log(req.target_id, 0.1)
    t_wait_ms = int((time.time() - t0) * 1000)
    debug_steps.append(f"wait window: {t_wait_ms} ms, hwnd={hwnd}")

    if hwnd:
        TARGET_MAP[req.target_id] = hwnd

    if CLOSE_HANDLE_AFTER_LAUNCH and final_pid:
        ok, msg = close_other_instance_handle(final_pid)
    else:
        ok, msg = (False, "Skipped")
    debug_steps.append(f"handle64: ok={ok} msg={msg}")

    if hwnd and POST_LAUNCH.get("enabled", True):
        threading.Thread(target=_do_post_launch, args=(req.target_id, hwnd), daemon=True).start()
        debug_steps.append(f"post_launch: started background seq={POST_LAUNCH.get('sequence','default')}")
    else:
        debug_steps.append("post_launch: disabled")

    return log_and_return(req.target_id, {
        "ok": True,
        "resolved": {"cmd": cmd, "cwd": cwd, "exe": cmd[0], "args": cmd[1:]},
        "pid": final_pid,
        "hwnd": hwnd,
        "handle_close": {"ok": ok, "msg": msg},
        "window_wait_ms": t_wait_ms,
        "steps": debug_steps,
    })

@app.post("/stop")
def stop(req: StopReq):
    pid = TARGET_PID.get(req.target_id)
    if not pid:
        return log_and_return(req.target_id, {"ok": False, "error": "no known pid for target"})
    try:
        p = psutil.Process(pid)
        hwnd = TARGET_MAP.get(req.target_id) or find_top_window_for_pid(pid)
        if hwnd and not req.force:
            win32api.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        else:
            if req.force:
                p.kill()
            else:
                p.terminate()
        return log_and_return(req.target_id, {"ok": True})
    except psutil.NoSuchProcess:
        return log_and_return(req.target_id, {"ok": True, "note": "process already gone"})
    except Exception as e:
        return log_and_return(req.target_id, {"ok": False, "error": str(e)})

@app.post("/focus")
def focus(req: FocusReq):
    hwnd = TARGET_MAP.get(req.target_id)
    if not hwnd:
        refresh_targets()
    hwnd = TARGET_MAP.get(req.target_id)
    if not hwnd:
        return log_and_return(req.target_id, {"ok": False, "error": "target not found"})
    ensure_restored_no_focus(hwnd)
    return log_and_return(req.target_id, {"ok": True})

@app.post("/click")
def click(req: ClickReq):
    hwnd = TARGET_MAP.get(req.target_id)
    cx, cy = int(req.cx), int(req.cy)
    w, h = get_client_size(hwnd)
    log_event(req.target_id, f"click client@{cx},{cy}")
    bg_mouse_click_client(hwnd, cx, cy, target_id=req.target_id)
    return JSONResponse({"ok": True, "clicked": {"client": [cx, cy], "wh": [w, h]}})

@app.post("/type")
def type_text(req: TypeReq):
    hwnd = TARGET_MAP.get(req.target_id)
    if not hwnd:
        refresh_targets()
        hwnd = TARGET_MAP.get(req.target_id)
        if not hwnd:
            return log_and_return(req.target_id, {"ok": False, "error": "target not found"})
    ensure_restored_no_focus(hwnd)
    for ch in req.text:
        code = ord(ch)
        if ch == "\n":
            ui_press_enter(hwnd, target_id=req.target_id)
        elif ch == "\t":
            bg_send_hotkey(hwnd, [win32con.VK_TAB], target_id=req.target_id)
        elif 32 <= code < 127:
            bg_send_char(hwnd, ch, target_id=req.target_id)
        else:
            vk = win32api.VkKeyScan(ch) & 0xFF
            bg_send_hotkey(hwnd, [vk], target_id=req.target_id)
        sleep_log(req.target_id, 0.005)
    return log_and_return(req.target_id, {"ok": True})

@app.post("/join_game")
def join_game(req: JoinReq):
    with JOIN_LOCK:
        steps = []
        hwnd = TARGET_MAP.get(req.target_id)
        if not hwnd:
            refresh_targets()
            hwnd = TARGET_MAP.get(req.target_id)
            if not hwnd:
                return log_and_return(req.target_id, {"ok": False, "error": "target not found"})
        ensure_restored_no_focus(hwnd)

        ui = load_uimap(UIMAP_PATH)
        required = ["GameNameBox"]
        for key in required:
            if key not in ui:
                return log_and_return(req.target_id, {"ok": False, "error": f"UiMap missing key: {key}"})

        cx, cy = uimap_point_client(hwnd, tuple(ui["GameNameBox"]))
        bg_mouse_click_client(hwnd, cx, cy, target_id=req.target_id)
        sleep_log(req.target_id, 0.05)
        steps.append(f"click GameNameBox client@{cx},{cy}")

        set_clipboard_text(req.game_name)
        bg_send_hotkey(hwnd, [win32con.VK_CONTROL, ord('A')], target_id=req.target_id)
        sleep_log(req.target_id, 0.5)
        bg_send_hotkey(hwnd, [win32con.VK_CONTROL, ord('V')], target_id=req.target_id)
        steps.append(f"copy paste game name: ({req.game_name})")
        
        '''
        for _ in range(20):
            bg_send_hotkey(hwnd, [win32con.VK_BACK], target_id=req.target_id)
            sleep_log(req.target_id, 0.005)
        steps.append("clear game name")

        for ch in req.game_name:
            bg_send_char(hwnd, ch, target_id=req.target_id)
            sleep_log(req.target_id, 0.005)
        steps.append(f"type game name ({len(req.game_name)} chars)")
        '''

        if (req.password or "") != "":
            bg_send_hotkey(hwnd, [win32con.VK_TAB], target_id=req.target_id)
            sleep_log(req.target_id, 0.5)  
            set_clipboard_text(req.password)
            bg_send_hotkey(hwnd, [win32con.VK_CONTROL, ord('A')], target_id=req.target_id)
            sleep_log(req.target_id, 0.5)
            bg_send_hotkey(hwnd, [win32con.VK_CONTROL, ord('V')], target_id=req.target_id)
            steps.append(f"copy paste game password: ({req.password})")

        ui_press_enter(hwnd, target_id=req.target_id)
        steps.append("press ENTER")
        return log_and_return(req.target_id, {"ok": True, "steps": steps})

@app.post("/leave_game")
def leave_game(req: LeaveReq):
    with JOIN_LOCK:
        steps = []
        hwnd = TARGET_MAP.get(req.target_id)
        if not hwnd:
            refresh_targets()
            hwnd = TARGET_MAP.get(req.target_id)
            if not hwnd:
                return log_and_return(req.target_id, {"ok": False, "error": "target not found"})
        ensure_restored_no_focus(hwnd)

        ui = load_uimap(UIMAP_PATH)
        ui_press_esc(hwnd, target_id=req.target_id)
        steps.append("press ESC")
        sleep_log(req.target_id, 0.15)

        if "LeaveButton" in ui:
            cx, cy = uimap_point_client(hwnd, tuple(ui["LeaveButton"]))
            cx, cy = (cx, cy)
            bg_mouse_click_client(hwnd, cx, cy, target_id=req.target_id)
            steps.append(f"click LeaveButton client@{cx},{cy}")
            sleep_log(req.target_id, 0.15)

        if "LeaveConfirm" in ui:
            cx, cy = uimap_point_client(hwnd, tuple(ui["LeaveConfirm"]))
            cx, cy = (cx, cy)
            bg_mouse_click_client(hwnd, cx, cy, target_id=req.target_id)
            steps.append(f"click LeaveConfirm client@{cx},{cy}")
        else:
            ui_press_enter(hwnd, target_id=req.target_id)
            steps.append("press ENTER")
        return log_and_return(req.target_id, {"ok": True, "steps": steps})

@app.post("/drain_logs")
def drain_logs(max_items: int = Body(200, embed=True)):
    out = []
    with LOGQ_LOCK:
        for _ in range(min(max_items, len(LOGQ))):
            out.append(LOGQ.popleft())
    return {"events": out}

# ============== Post-Launch sequence ==============

def _do_post_launch(target_id: str, hwnd: int):
    log_event(target_id, "post: start")
    if not POST_LAUNCH.get("enabled", True):
        return
    seq = POST_LAUNCH.get("sequence", "default")
    wait_for_start_up = POST_LAUNCH.get("wait_for_start_up", "5.0")
    wait_for_title_load_up = POST_LAUNCH.get("wait_for_title_load_up", "55.0")
    wait_for_connect_to_server = POST_LAUNCH.get("wait_for_connect_to_server", "45.0")
    try:
        if seq == "default":
            sleep_log(target_id, wait_for_start_up)           
            ui_press_space(hwnd, target_id)
            sleep_log(target_id, 2.0)
            ui_press_space(hwnd, target_id)
            sleep_log(target_id, wait_for_title_load_up)
            ui_press_space(hwnd, target_id)
            sleep_log(target_id, wait_for_connect_to_server)
            log_event(target_id, "post: calling goto_lobby")
            ret = _do_goto_lobby(target_id, hwnd)
            log_event(target_id, f"post: goto_lobby result={ret}")
    except Exception as e:
        log_event(target_id, f"post: error {e}")
    finally:
        log_event(target_id, "post: end")

# ============== Main ==============

def main():
    parser = argparse.ArgumentParser(description="D2R Worker (absolute clicks, debug logs, non-blocking post-launch)")
    parser.add_argument("--config", type=str, default="worker/config.json")
    parser.add_argument("--name", type=str, help="Worker name override")
    parser.add_argument("--port", type=int, help="Port override")
    parser.add_argument("--title", type=str, help="Window title match override")
    args = parser.parse_args()

    cfg_path = Path(args.config)
    if cfg_path.exists():
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    else:
        cfg = {}

    global WORKER_NAME, WINDOW_TITLE_MATCH, UIMAP_PATH, LAUNCHERS, POST_LAUNCH, CLOSE_HANDLE_AFTER_LAUNCH
    WORKER_NAME = args.name or cfg.get("worker_name", WORKER_NAME)
    WINDOW_TITLE_MATCH = args.title or cfg.get("window_title_match", WINDOW_TITLE_MATCH)
    UIMAP_PATH = cfg.get("uimap_path", UIMAP_PATH)
    LAUNCHERS = cfg.get("launchers", {})
    POST_LAUNCH = cfg.get("post_launch", POST_LAUNCH)
    CLOSE_HANDLE_AFTER_LAUNCH = bool(cfg.get("close_handle_after_launch", CLOSE_HANDLE_AFTER_LAUNCH))

    global LOCK_ASPECT, ASPECT_W, ASPECT_H, ASPECT_ANCHOR, ASPECT_MODE
    LOCK_ASPECT = bool(cfg.get("lock_aspect", LOCK_ASPECT))
    ASPECT_W = int(cfg.get("aspect_w", ASPECT_W))
    ASPECT_H = int(cfg.get("aspect_h", ASPECT_H))
    ASPECT_ANCHOR = (cfg.get("aspect_anchor", ASPECT_ANCHOR) or "topleft")
    ASPECT_MODE = (cfg.get("aspect_mode", ASPECT_MODE) or "auto")

    port = args.port or int(cfg.get("port", 5001))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")

if __name__ == "__main__":
    main()
