"""
capturer.py - Screen capture and slide-change detection logic.

Uses mss for fast screen capture and Pillow for pixel-diff comparison.
A slide is considered "changed" when the mean pixel difference between
the previous and current screenshots exceeds a configurable threshold.
"""

import threading
import time
from pathlib import Path
from typing import Callable, List, Optional

import mss
import mss.tools
from PIL import Image, ImageChops, ImageStat
from pynput import keyboard as _kb


def get_monitor_under_mouse() -> dict:
    """Return the mss monitor dict for the screen where the mouse cursor is."""
    import pyautogui
    mx, my = pyautogui.position()
    with mss.mss() as sct:
        for monitor in sct.monitors[1:]:
            if (monitor["left"] <= mx < monitor["left"] + monitor["width"] and
                    monitor["top"] <= my < monitor["top"] + monitor["height"]):
                return dict(monitor)
        # Fallback: primary monitor
        return dict(sct.monitors[1])


def capture_region(region: dict) -> Image.Image:
    """Capture the given screen region and return as a PIL Image."""
    monitor = {
        "left": region["left"],
        "top": region["top"],
        "width": region["width"],
        "height": region["height"],
    }
    with mss.mss() as sct:
        raw = sct.grab(monitor)
        return Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")


def is_end_screen(img: Image.Image, brightness_threshold: float = 8.0) -> bool:
    """
    Return True if the image is mostly black (end-of-slideshow screen).
    Checks mean brightness across all pixels. Threshold is 0-255 scale.
    """
    stat = ImageStat.Stat(img)
    mean_brightness = sum(stat.mean) / len(stat.mean)
    return mean_brightness < brightness_threshold


def images_are_different(img_a: Image.Image, img_b: Image.Image, threshold: float = 1.0) -> bool:
    """
    Compare two images. Returns True if the mean channel difference
    as a percentage exceeds `threshold` (0-100).
    """
    if img_a.size != img_b.size:
        return True
    diff = ImageChops.difference(img_a, img_b)
    stat = ImageStat.Stat(diff)
    mean_diff = sum(stat.mean) / len(stat.mean)
    # mean_diff is 0-255; convert to percentage
    diff_pct = (mean_diff / 255.0) * 100
    return diff_pct > threshold


def wait_for_change_then_stable(
    region: dict,
    reference_img: Image.Image,
    change_threshold: float = 0.5,
    stable_threshold: float = 0.2,
    stable_duration: float = 0.4,
    poll_interval: float = 0.08,
    timeout: float = 8.0,
) -> Image.Image:
    """
    1. Poll until the screen differs from reference_img (slide started changing).
    2. Then poll until the screen stops changing (animation finished).
    3. Return the stable image.

    This replaces a fixed delay — reacts instantly and captures as soon as stable.
    """
    deadline = time.time() + timeout

    # Phase 1: wait for ANY change from reference
    while time.time() < deadline:
        current = capture_region(region)
        if images_are_different(reference_img, current, threshold=change_threshold):
            break
        time.sleep(poll_interval)
    else:
        return capture_region(region)

    # Phase 2: wait for screen to stop moving (animation done)
    last_img = capture_region(region)
    stable_since = time.time()
    while time.time() < deadline:
        time.sleep(poll_interval)
        current = capture_region(region)
        if images_are_different(last_img, current, threshold=stable_threshold):
            last_img = current
            stable_since = time.time()
        else:
            if time.time() - stable_since >= stable_duration:
                return current
    return capture_region(region)


def refine_slide_region(base_region: dict, skip_top: int = 55, skip_bottom: int = 45) -> dict:
    """
    Fine-tune a viewport region to exclude browser/viewer toolbars and margins.

    Strategy:
    1. Skip the top toolbar (skip_top px — Office Online toolbar for PDF; 0 for PPTX).
    2. Sample the left/right edges to detect the viewer background color.
    3a. Light background (brightness > 180): find rows/cols with >40% bright
        pixels — the white slide content stands out from the gray bg.
    3b. Dark background (brightness ≤ 180): find rows/cols where pixels differ
        from the background color by > 20 — slide content is more colorful
        than the uniform dark margin.
    4. Trim bottom nav bar (skip_bottom px — PDF nav bar; 0 for PPTX).

    Falls back to base_region if detection fails or the area is too small.
    """
    import numpy as np

    SKIP_TOP    = skip_top    # Office Online toolbar (PDF=55, PPTX=0)
    SKIP_BOTTOM = skip_bottom # PDF navigation bar (PDF=45, PPTX=0)
    MIN_FRAC    = 0.30        # result must be ≥30% of original

    try:
        img = capture_region(base_region)
        arr = np.array(img.convert("RGB"), dtype=np.float32)
        h, w, _ = arr.shape
        
        PAD = 15  # Padding to keep around content (avoid cropping too tight)

        # Sample left/right edges (middle portion, below toolbar)
        mid_y0 = SKIP_TOP + int((h - SKIP_TOP) * 0.2)
        mid_y1 = SKIP_TOP + int((h - SKIP_TOP) * 0.8)
        left_mean  = arr[mid_y0:mid_y1, :20,  :].mean(axis=(0, 1))  # [R,G,B]
        right_mean = arr[mid_y0:mid_y1, -20:, :].mean(axis=(0, 1))

        bg_brightness = (left_mean.mean() + right_mean.mean()) / 2
        print(f"      [Refine] bg_brightness={bg_brightness:.1f} (Light > 180)")

        if bg_brightness > 180:
            # ── Light background: brightness-based ────────────────────────────
            gray = arr.mean(axis=2)
            row_frac = (gray > 230).mean(axis=1)
            col_frac = (gray > 230).mean(axis=0)
            bright_rows = np.where(row_frac > 0.40)[0]
            bright_cols = np.where(col_frac > 0.40)[0]
            if len(bright_rows) == 0 or len(bright_cols) == 0:
                print("      [Refine] No bright content found")
                return base_region
            first_row = max(0, int(bright_rows[0]) - PAD)
            last_row  = min(h, int(bright_rows[-1]) + PAD)
            first_col = max(0, int(bright_cols[0]) - PAD)
            last_col  = min(w, int(bright_cols[-1]) + PAD)
        else:
            # ── Dark background: color-diff based ─────────────────────────────
            bg_color = (left_mean + right_mean) / 2
            diff = np.abs(arr - bg_color).max(axis=2)
            is_content = diff > 20

            col_content = is_content[SKIP_TOP:].mean(axis=0)
            row_content = is_content.mean(axis=1)

            content_cols = np.where(col_content > 0.10)[0]
            content_rows = np.where((row_content > 0.10) &
                                    (np.arange(h) >= SKIP_TOP))[0]

            if len(content_rows) == 0 or len(content_cols) == 0:
                print("      [Refine] No diff content found")
                return base_region
            first_row = max(0, int(content_rows[0]) - PAD)
            last_row  = min(h, int(content_rows[-1]) + PAD)
            first_col = max(0, int(content_cols[0]) - PAD)
            last_col  = min(w, int(content_cols[-1]) + PAD)

        # Trim bottom nav bar
        last_row = max(first_row, last_row - SKIP_BOTTOM)

        new_h = last_row - first_row
        new_w = last_col - first_col
        
        print(f"      [Refine] Crop: top={first_row}, bottom={h-last_row}, left={first_col}, right={w-last_col}")

        if new_h < h * MIN_FRAC or new_w < w * MIN_FRAC:
            print(f"      [Refine] Result too small ({new_w}x{new_h}), fallback")
            return base_region

        return {
            "left":   base_region["left"]  + first_col,
            "top":    base_region["top"]   + first_row,
            "width":  new_w,
            "height": new_h,
        }
    except Exception as e:
        print(f"      [Refine] Error: {e}")
        return base_region


def save_slide_image(img: Image.Image, output_dir: Path, index: int) -> Path:
    """Save a slide image as slide_NNN.png and return the path."""
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = output_dir / f"slide_{index:03d}.png"
    img.save(filename, "PNG")
    return filename


def run_capture_session(
    region: dict,
    output_dir: Path,
    navigate_fn: Callable,
    delay: float = 0.0,
    max_slides: int = 200,
    diff_threshold: float = 1.0,
    same_count_limit: int = 3,
    exact_total: Optional[int] = None,
) -> List[Path]:
    """
    Main capture loop:
      1. Capture + save current slide.
      2. Press next.
      3. Reactively wait: detect change → wait for stable → capture.
      4. Repeat until no change for same_count_limit consecutive attempts.

    `delay` is an optional extra pause after pressing next (default 0 = reactive only).
    """
    saved_paths: List[Path] = []
    consecutive_same = 0
    slide_index = 1

    print(f"\nCapturing into: {output_dir}")
    print("Ctrl+C to stop early.\n")

    # Lắng nghe phím ESC để dừng sớm
    _stop_flag = threading.Event()

    def _on_press(key):
        if key == _kb.Key.esc:
            _stop_flag.set()
            return False  # dừng listener

    _listener = _kb.Listener(on_press=_on_press)
    _listener.start()
    print("  (Bấm ESC bất kỳ lúc nào để dừng)\n")

    first_slide_img = capture_region(region)
    last_saved_img = first_slide_img
    path = save_slide_image(first_slide_img, output_dir, slide_index)
    saved_paths.append(path)
    print(f"  Saved {path.name}")

    try:
        while slide_index < max_slides and not _stop_flag.is_set():
            # In ra log bấm Next rõ ràng, xuống dòng để không bị che
            print(f"  → Đang bấm Next (lần thử {consecutive_same + 1})...")
            navigate_fn()
            time.sleep(0.8 + delay)

            current_img = wait_for_change_then_stable(
                region,
                reference_img=last_saved_img,
                timeout=3.0,
            )

            # Phát hiện màn hình kết thúc (màn hình đen)
            if is_end_screen(current_img):
                print("\nPhát hiện màn hình kết thúc → nhấn ESC thoát presentation.")
                import pyautogui as _pg
                _pg.press("escape")
                time.sleep(1.2)  # chờ thoát presentation mode
                break  # tab closing is handled by --auto-close in main.py / crawler

            # Phát hiện loop: nếu ảnh giống slide_001 → nhấn ESC thoát + dừng
            if slide_index >= 3 and not images_are_different(first_slide_img, current_img, threshold=0.5):
                print("\nPhát hiện loop về slide đầu → nhấn ESC thoát presentation.")
                import pyautogui as _pg
                _pg.press("escape")
                break

            if not images_are_different(last_saved_img, current_img, threshold=0.1):
                consecutive_same += 1
                print(f"  Không thay đổi ({consecutive_same}/{same_count_limit})")
                if consecutive_same >= same_count_limit:
                    print("\nHết slide.")
                    break
            else:
                slide_index += 1
                path = save_slide_image(current_img, output_dir, slide_index)
                saved_paths.append(path)
                last_saved_img = current_img
                consecutive_same = 0
                print(f"  Saved {path.name}")
                if exact_total and slide_index >= exact_total:
                    print(f"\nĐã chụp đủ {exact_total} slide.")
                    break
    except KeyboardInterrupt:
        print("\nDừng sớm (Ctrl+C).")
    finally:
        _listener.stop()

    if _stop_flag.is_set():
        print("\nDừng sớm (ESC).")

    print(f"\nTổng số slide: {len(saved_paths)}")
    return saved_paths
