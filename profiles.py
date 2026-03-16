"""
profiles.py - Save and load named viewer profiles.

Each profile stores:
  - region: {left, top, width, height} of the slide content area (no browser chrome)
  - nav: {x, y} of the next-page button (or null for keyboard navigation)
  - key: keyboard key to use when nav is null (default: "right")

Profiles are saved to profiles.json next to this file.
"""

import json
import time
from pathlib import Path
from typing import Optional

import pyautogui

PROFILES_FILE = Path(__file__).parent / "profiles.json"


def load_profiles() -> dict:
    if PROFILES_FILE.exists():
        with open(PROFILES_FILE) as f:
            return json.load(f)
    return {}


def save_profiles(profiles: dict) -> None:
    with open(PROFILES_FILE, "w") as f:
        json.dump(profiles, f, indent=2)


def get_profile(name: str) -> Optional[dict]:
    return load_profiles().get(name)


def list_profiles() -> list:
    return list(load_profiles().keys())


def _countdown_position(label: str, seconds: int = 8) -> tuple:
    """Show live mouse position countdown, return final (x, y)."""
    print(f"\n  Bấm Enter khi sẵn sàng di chuột đến {label}: ", end="", flush=True)
    input()
    print(f"  Di chuột đến {label} ({seconds} giây):")
    for i in range(seconds, 0, -1):
        pos = pyautogui.position()
        print(f"\r    [{i}s] x={pos.x:5d}, y={pos.y:5d}   ", end="", flush=True)
        time.sleep(1)
    final = pyautogui.position()
    print(f"\r    [OK] {label}: x={final.x}, y={final.y}          ")
    return final.x, final.y


def setup_profile(name: str) -> dict:
    """
    Interactive setup for a named viewer profile.
    Returns the profile dict.
    """
    print(f"\n{'='*55}")
    print(f"  SETUP PROFILE: {name}")
    print(f"{'='*55}")

    # ── Bước 1: Region (tùy chọn) ────────────────────────────────────────────
    print("\nBước 1 — Vùng crop slide (bỏ qua nếu chụp toàn màn hình)")
    use_region = input("  Cần crop vùng slide cụ thể? (y/n) [mặc định: n]: ").strip().lower() == "y"

    region = None
    if use_region:
        x1, y1 = _countdown_position("GÓC TRÊN-TRÁI")
        x2, y2 = _countdown_position("GÓC DƯỚI-PHẢI")
        region = {
            "left":   min(x1, x2),
            "top":    min(y1, y2),
            "width":  abs(x2 - x1),
            "height": abs(y2 - y1),
        }
        print(f"  Slide region: {region['width']}x{region['height']} tại ({region['left']}, {region['top']})")
    else:
        print("  Sẽ chụp toàn màn hình nơi con trỏ chuột đang đứng.")

    # ── Bước 2: Nav button ────────────────────────────────────────────────────
    print("\nBước 2 — Nút NEXT (chuyển trang)")
    use_click = input("  Dùng click chuột cho nút next? (y/n) [mặc định: n = phím bàn phím]: ").strip().lower() == "y"

    nav = None
    key = "right"
    if use_click:
        nx, ny = _countdown_position("nút NEXT")
        nav = {"x": nx, "y": ny}
    else:
        key = input("  Nhập tên phím (right/down/pagedown) [mặc định: right]: ").strip() or "right"

    # ── Bước 3: Nút Trình bày (tùy chọn) ─────────────────────────────────────
    print("\nBước 3 — Nút TRÌNH BÀY / Slideshow (tùy chọn)")
    use_present = input("  Lưu vị trí nút Trình bày để tự động click? (y/n): ").strip().lower() == "y"

    present_btn = None
    if use_present:
        px, py = _countdown_position("nút TRÌNH BÀY")
        present_btn = {"x": px, "y": py}

    profile = {
        "region":      region,
        "nav":         nav,
        "key":         key,
        "present_btn": present_btn,
    }

    profiles = load_profiles()
    profiles[name] = profile
    save_profiles(profiles)
    print(f"\nProfile '{name}' đã lưu vào profiles.json")
    return profile
