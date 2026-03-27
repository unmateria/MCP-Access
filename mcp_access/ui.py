"""
Screenshot capture and UI automation (click, type).
"""

import ctypes
import os
import tempfile
import threading
import time
from datetime import datetime
from typing import Any, Optional

from .core import _Session, log


# ---------------------------------------------------------------------------
# Window capture via PrintWindow API
# ---------------------------------------------------------------------------

def _capture_window(hwnd: int, max_width: int = 1920) -> tuple:
    """
    Capture an Access window using PrintWindow API.
    Returns (PIL.Image, original_width, original_height).
    """
    import win32gui
    import win32ui
    from PIL import Image

    # Get window dimensions
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bottom - top
    if w <= 0 or h <= 0:
        raise RuntimeError(f"Window has invalid dimensions: {w}x{h}")

    # Create device context and bitmap
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, w, h)
    save_dc.SelectObject(bitmap)

    # Capture — PW_RENDERFULLCONTENT = 2 (works even if partially obscured)
    ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

    # Convert to PIL Image
    bmpinfo = bitmap.GetInfo()
    bmpstr = bitmap.GetBitmapBits(True)
    img = Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                           bmpstr, "raw", "BGRX", 0, 1)

    # Cleanup GDI resources
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    win32gui.DeleteObject(bitmap.GetHandle())

    original_w, original_h = w, h

    # Resize if wider than max_width
    if w > max_width:
        ratio = max_width / w
        new_h = int(h * ratio)
        img = img.resize((max_width, new_h), Image.LANCZOS)

    return img, original_w, original_h


# ---------------------------------------------------------------------------
# Screenshot
# ---------------------------------------------------------------------------

def ac_screenshot(
    db_path: str,
    object_type: str = "",
    object_name: str = "",
    output_path: str = "",
    wait_ms: int = 300,
    max_width: int = 1920,
    open_timeout_sec: int = 30,
) -> dict:
    """Capture the Access window as PNG. Optionally opens a form/report first.

    NOTA: Timer events de Access NO se disparan durante la captura (no hay
    Windows message pump). Si el form usa Form_Timer para inicializacion
    (ej: WebBrowser navigate), abrir el form manualmente antes, o usar
    access_run_vba para forzar la inicializacion.

    open_timeout_sec: segundos maximos esperando que OpenForm complete (default 30).
    Si el Form_Load/Open tarda mas (ej: OpenRecordset lento), se envia ESC para
    cancelar la operacion y se lanza TimeoutError. Evita bloqueos de 40+ minutos.
    """
    import win32gui
    import win32api
    import win32con

    app = _Session.connect(db_path)
    object_opened = False

    # Open form/report if requested
    if object_type and object_name:
        ot = object_type.lower()
        if ot not in ("form", "report"):
            raise ValueError(f"object_type must be 'form' or 'report', got '{object_type}'")

        # Get hwnd before OpenForm blocks (needed by cancel thread)
        _h = app.hWndAccessApp
        _hwnd = int(_h() if callable(_h) else _h)

        # Background thread: send ESC after timeout to cancel hanging Load events
        _done = threading.Event()
        _timed_out = threading.Event()

        def _cancel_if_hung():
            if not _done.wait(open_timeout_sec):
                _timed_out.set()
                log.warning(
                    "OpenForm '%s' timeout after %ds — sending ESC to cancel",
                    object_name, open_timeout_sec,
                )
                win32api.PostMessage(_hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
                win32api.PostMessage(_hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)

        _t = threading.Thread(target=_cancel_if_hung, daemon=True)
        _t.start()
        try:
            if ot == "form":
                app.DoCmd.OpenForm(object_name, 0)  # acNormal
            else:
                app.DoCmd.OpenReport(object_name, 2)  # acPreview
        finally:
            _done.set()

        if _timed_out.is_set():
            raise TimeoutError(
                f"OpenForm '{object_name}' did not complete within {open_timeout_sec}s. "
                "Form_Load event may have a slow/blocked OpenRecordset. "
                "ESC was sent to cancel. Increase open_timeout_sec if the form is intentionally slow."
            )
        object_opened = True

    if wait_ms > 0:
        import pythoncom
        _deadline = time.time() + wait_ms / 1000.0
        while time.time() < _deadline:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.015)  # ~60 Hz, prevent busy-wait

    _h = app.hWndAccessApp
    hwnd = int(_h() if callable(_h) else _h)

    # Restore if minimized
    if ctypes.windll.user32.IsIconic(hwnd):
        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        time.sleep(0.3)

    img, orig_w, orig_h = _capture_window(hwnd, max_width)

    # Close the object we opened (leave it clean)
    if object_opened:
        ot = object_type.lower()
        try:
            ac_type_code = 2 if ot == "form" else 3  # acForm / acReport
            app.DoCmd.Close(ac_type_code, object_name, 1)  # acSaveNo
        except Exception as e:
            log.warning("Could not close %s %s: %s", object_type, object_name, e)

    # Determine output path
    if not output_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(tempfile.gettempdir(), f"access_screenshot_{ts}.png")

    # Ensure directory exists
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    img.save(output_path, "PNG")
    file_size = os.path.getsize(output_path)

    return {
        "path": output_path,
        "width": img.width,
        "height": img.height,
        "original_width": orig_w,
        "original_height": orig_h,
        "file_size": file_size,
        "object_opened": f"{object_type}:{object_name}" if object_opened else None,
    }


# ---------------------------------------------------------------------------
# UI Click
# ---------------------------------------------------------------------------

def ac_ui_click(
    db_path: str,
    x: int,
    y: int,
    image_width: int,
    click_type: str = "left",
    wait_after_ms: int = 200,
) -> dict:
    """Click at image coordinates on the Access window."""
    import win32api
    import win32gui

    app = _Session.connect(db_path)
    _h = app.hWndAccessApp
    hwnd = int(_h() if callable(_h) else _h)

    # Bring to foreground
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(0.05)

    # Get window rect for coordinate scaling
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win_w = right - left
    win_h = bottom - top

    # Scale image coords -> screen coords
    scale = win_w / image_width
    screen_x = int(left + x * scale)
    screen_y = int(top + y * scale)

    # Move cursor and click
    win32api.SetCursorPos((screen_x, screen_y))
    time.sleep(0.02)

    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    MOUSEEVENTF_RIGHTDOWN = 0x0008
    MOUSEEVENTF_RIGHTUP = 0x0010

    ct = click_type.lower()
    if ct == "left":
        win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
    elif ct == "double":
        win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
    elif ct == "right":
        win32api.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0)
        win32api.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0)
    else:
        raise ValueError(f"click_type must be 'left', 'double', or 'right', got '{click_type}'")

    if wait_after_ms > 0:
        time.sleep(wait_after_ms / 1000.0)

    return {
        "clicked_screen_x": screen_x,
        "clicked_screen_y": screen_y,
        "click_type": ct,
    }


# ---------------------------------------------------------------------------
# UI Type / Send keys
# ---------------------------------------------------------------------------

def ac_ui_type(
    db_path: str,
    text: str = "",
    key: str = "",
    modifiers: str = "",
    wait_after_ms: int = 100,
) -> dict:
    """Type text or send keyboard shortcuts to the Access window."""
    import win32api
    import win32gui

    if not text and not key:
        raise ValueError("Must provide either 'text' or 'key'")

    app = _Session.connect(db_path)
    _h = app.hWndAccessApp
    hwnd = int(_h() if callable(_h) else _h)

    # Bring to foreground
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(0.05)

    VK_MAP = {
        "enter": 0x0D, "tab": 0x09, "escape": 0x1B, "backspace": 0x08,
        "delete": 0x2E, "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
        "home": 0x24, "end": 0x23, "space": 0x20,
        "pageup": 0x21, "pagedown": 0x22,
        "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73,
        "f5": 0x74, "f6": 0x75, "f7": 0x76, "f8": 0x77,
        "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
    }
    MOD_MAP = {
        "ctrl": 0x11, "shift": 0x10, "alt": 0x12,
    }

    result_desc = ""

    if text:
        # Type each character using WM_CHAR
        WM_CHAR = 0x0102
        for ch in text:
            win32api.SendMessage(hwnd, WM_CHAR, ord(ch), 0)
            time.sleep(0.01)
        result_desc = f"typed: {text}"

    if key:
        vk = VK_MAP.get(key.lower())
        if vk is None:
            # Try single letter/digit as VkKeyScan
            if len(key) == 1:
                vk = ctypes.windll.user32.VkKeyScanW(ord(key)) & 0xFF
            else:
                raise ValueError(f"Unknown key: '{key}'. Valid: {list(VK_MAP.keys())}")

        # Press modifiers
        mod_keys = []
        if modifiers:
            for mod in modifiers.lower().split("+"):
                mod = mod.strip()
                mvk = MOD_MAP.get(mod)
                if mvk is None:
                    raise ValueError(f"Unknown modifier: '{mod}'. Valid: ctrl, shift, alt")
                mod_keys.append(mvk)
                win32api.keybd_event(mvk, 0, 0, 0)  # key down
                time.sleep(0.01)

        # Press and release the key
        win32api.keybd_event(vk, 0, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(vk, 0, 2, 0)  # KEYEVENTF_KEYUP = 2

        # Release modifiers (reverse order)
        for mvk in reversed(mod_keys):
            win32api.keybd_event(mvk, 0, 2, 0)

        mod_str = f"{modifiers}+" if modifiers else ""
        result_desc = f"key: {mod_str}{key}"

    if wait_after_ms > 0:
        time.sleep(wait_after_ms / 1000.0)

    return {
        "action": result_desc,
        "modifiers": modifiers if modifiers else None,
    }
