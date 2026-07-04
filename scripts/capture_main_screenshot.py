from __future__ import annotations

import sys
import time
import ctypes
import ctypes.wintypes
from pathlib import Path

from PIL import ImageGrab


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from tkinter import Tk

from vhlookup_app.tk_main import LocalApp


try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass


def main() -> int:
    output_path = ROOT / "docs" / "assets" / "vhlookup-local-main.jpg"
    root = Tk()
    LocalApp(root)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    width = min(1200, max(900, screen_width - 24))
    height = min(980, max(680, screen_height - 70))
    root.geometry(f"{width}x{height}+10+10")
    root.update_idletasks()
    root.update()
    root.lift()
    try:
        root.attributes("-topmost", True)
        root.update()
        time.sleep(0.8)
        root.attributes("-topmost", False)
        root.update()
    except Exception:
        pass
    rect = ctypes.wintypes.RECT()
    hwnd = root.winfo_id()
    if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        bbox = (rect.left, rect.top, rect.right, rect.bottom)
    else:
        x = root.winfo_rootx()
        y = root.winfo_rooty()
        width = root.winfo_width()
        height = root.winfo_height()
        bbox = (x, y, x + width, y + height)
    image = ImageGrab.grab(bbox=bbox)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path, quality=92)
    print(output_path)
    root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
