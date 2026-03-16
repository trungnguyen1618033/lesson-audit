"""
navigator.py - Slide navigation via keyboard or mouse click.

macOS requirement: Terminal must be granted Accessibility access in
System Settings > Privacy & Security > Accessibility.
"""

import time
from typing import Optional

import pyautogui

pyautogui.PAUSE = 0.05

# Global click position for click-based navigation
_next_btn_pos: Optional[tuple] = None

# Global key for keyboard-based navigation (default: right arrow)
_nav_key: str = "right"


def set_next_button_position(x: int, y: int) -> None:
    """Store the screen position of the 'next page' button."""
    global _next_btn_pos
    _next_btn_pos = (x, y)


def set_nav_key(key: str) -> None:
    """Set the keyboard key used to advance slides (e.g. 'right', 'pagedown', 'down')."""
    global _nav_key
    _nav_key = key


def capture_next_button_position(seconds: int = 8) -> tuple:
    """
    Countdown while user hovers over the 'next page' (v) button.
    Returns (x, y) of the button when countdown ends.
    """
    print("\nDi chuột đến nút chuyển trang tiếp theo (nút v ▼ trên toolbar):")
    print("Bấm Enter khi sẵn sàng: ", end="", flush=True)
    input()
    print(f"Di chuột đến nút v ngay bây giờ (có {seconds} giây):")
    for i in range(seconds, 0, -1):
        pos = pyautogui.position()
        print(f"\r  [{i}s] Chuột đang ở: x={pos.x:5d}, y={pos.y:5d}   ", end="", flush=True)
        time.sleep(1)
    final = pyautogui.position()
    print(f"\r  [OK] Đã lưu vị trí nút v: x={final.x}, y={final.y}          ")
    set_next_button_position(final.x, final.y)
    return final.x, final.y


def press_next() -> None:
    """Advance to next slide: click button if position set, else use configured key."""
    if _next_btn_pos:
        pyautogui.click(_next_btn_pos[0], _next_btn_pos[1])
    else:
        pyautogui.press(_nav_key)


def press_prev() -> None:
    pyautogui.press("pageup")


def press_home() -> None:
    pyautogui.hotkey("command", "home")


def click_slide_area(region: dict) -> None:
    cx = region["left"] + region["width"] // 2
    cy = region["top"] + region["height"] // 2
    pyautogui.click(cx, cy)
    time.sleep(0.3)


def focus_and_home(region: dict) -> None:
    click_slide_area(region)
    time.sleep(0.3)
    press_home()
    time.sleep(0.5)
