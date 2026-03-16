"""
video_capture.py — Capture video (MP4) files from SharePoint.

Two strategies:
  1. Network interception: Intercept video stream URL via Playwright, download directly.
  2. Screen recording: Record screen + system audio with ffmpeg (fallback).

macOS audio note: System audio capture requires BlackHole virtual audio device.
  Install: brew install blackhole-2ch
  Then create a Multi-Output Device in Audio MIDI Setup.
"""

import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

from playwright.sync_api import Page


# ── Strategy 1: Network interception ─────────────────────────────────────────

def _intercept_video_url(page: Page, video_page_url: str, timeout: float = 15.0) -> Optional[str]:
    """
    Open a SharePoint video page and intercept network requests to find the
    actual video stream URL (.mp4 / video content-type).
    Returns the direct video URL or None.
    """
    captured_urls: list[str] = []

    def _on_response(response):
        url = response.url
        ct = response.headers.get("content-type", "")
        # Look for video content or .mp4 URLs
        if "video/" in ct or ".mp4" in urlparse(url).path.lower():
            captured_urls.append(url)
        # SharePoint streaming manifests
        if "videomanifest" in url.lower() or "getvideostream" in url.lower():
            captured_urls.append(url)

    page.on("response", _on_response)

    try:
        page.goto(video_page_url, timeout=30000, wait_until="domcontentloaded")
        time.sleep(5.0)

        # Try to extract video src from the DOM (video tag or iframe)
        dom_src = page.evaluate("""
        () => {
            // Direct video element
            const v = document.querySelector('video');
            if (v && v.src) return v.src;
            if (v) {
                const source = v.querySelector('source');
                if (source && source.src) return source.src;
            }
            // Check iframes
            try {
                const frames = document.querySelectorAll('iframe');
                for (const f of frames) {
                    try {
                        const fv = f.contentDocument?.querySelector('video');
                        if (fv && fv.src) return fv.src;
                    } catch(e) {}
                }
            } catch(e) {}
            return null;
        }
        """)
        if dom_src:
            captured_urls.insert(0, dom_src)

        # Also check for video in iframes via Playwright frames
        for frame in page.frames:
            try:
                frame_src = frame.evaluate("""
                () => {
                    const v = document.querySelector('video');
                    if (v && v.src) return v.src;
                    if (v) {
                        const source = v.querySelector('source');
                        if (source && source.src) return source.src;
                    }
                    return null;
                }
                """)
                if frame_src:
                    captured_urls.insert(0, frame_src)
            except Exception:
                pass

        # Click play if video is paused
        try:
            # 1. Native click the center of the video area (viewport is shifted right by sidebar, so click width // 3)
            page.mouse.click(page.viewport_size["width"] // 3, page.viewport_size["height"] // 2)
            time.sleep(1.0)
            
            # 2. Look for any Play buttons in the DOM and click them
            play_buttons = page.locator('button, [role="button"]').filter(has_text="Play").all()
            for pb in play_buttons:
                try:
                    pb.click(timeout=1000)
                except Exception:
                    pass
            time.sleep(1.0)

            # 3. Direct JS manipulation
            for frame in page.frames:
                try:
                    frame.evaluate("""
                    () => {
                        const v = document.querySelector('video');
                        if (v && v.paused) v.play();
                    }
                    """)
                except Exception:
                    pass
                    
            # 4. OS-level click using pyautogui (ultimate fallback for stubborn iframes)
            import pyautogui
            from crawler import _get_viewport_screen_region
            vp = _get_viewport_screen_region(page)
            if vp:
                cx = vp["left"] + vp["width"] // 3
                cy = vp["top"] + vp["height"] // 2
                pyautogui.click(cx, cy)
                time.sleep(0.5)
        except Exception:
            pass

        # Wait for video requests to appear
        deadline = time.time() + timeout
        while time.time() < deadline and not captured_urls:
            time.sleep(1.0)

    finally:
        try:
            page.remove_listener("response", _on_response)
        except Exception:
            pass

    # Prefer .mp4 URLs, then video content-type URLs
    mp4_urls = [u for u in captured_urls if ".mp4" in urlparse(u).path.lower()]
    if mp4_urls:
        return mp4_urls[0]
    return captured_urls[0] if captured_urls else None


def download_video(url: str, output_path: Path, cookies: Optional[dict] = None) -> bool:
    """Download video from URL to output_path. Returns True on success."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    print(f"    Downloading video: {output_path.name}")

    try:
        req = urllib.request.Request(url)
        if cookies:
            cookie_str = "; ".join(f"{k}={v}" for k, v in cookies.items())
            req.add_header("Cookie", cookie_str)

        with urllib.request.urlopen(req, timeout=300) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            with open(output_path, "wb") as f:
                while True:
                    chunk = resp.read(1024 * 1024)  # 1MB chunks
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = downloaded * 100 // total
                        mb = downloaded / (1024 * 1024)
                        print(f"\r    Downloaded: {mb:.1f}MB ({pct}%)", end="", flush=True)

        print()
        size_mb = output_path.stat().st_size / (1024 * 1024)
        print(f"    Saved: {output_path} ({size_mb:.1f}MB)")
        return True
    except Exception as e:
        print(f"    [ERROR] Download failed: {e}")
        if output_path.exists():
            output_path.unlink()
        return False


def download_via_playwright(page: Page, url: str, output_path: Path) -> bool:
    """Download video using Playwright's browser context (preserves auth cookies)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    print(f"    Downloading video via browser: {output_path.name}")

    try:
        resp = page.request.get(url, timeout=300000)
        if resp.ok:
            output_path.write_bytes(resp.body())
            size_mb = output_path.stat().st_size / (1024 * 1024)
            print(f"    Saved: {output_path} ({size_mb:.1f}MB)")
            return True
        else:
            print(f"    [ERROR] HTTP {resp.status}: {resp.status_text}")
            return False
    except Exception as e:
        print(f"    [ERROR] Browser download failed: {e}")
        return False


def download_stream_ffmpeg(page: Page, url: str, output_path: Path) -> bool:
    """Download HLS/DASH streaming video using ffmpeg with browser cookies."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    print(f"    [INFO] Bắt đầu tải luồng Streaming (HLS/DASH) qua ffmpeg...")
    
    try:
        # Lấy cookies từ Playwright context để ffmpeg có thể vượt qua xác thực
        cookies = page.context.cookies()
        cookie_str = "; ".join(f"{c['name']}={c['value']}" for c in cookies)
        
        # Gọi ffmpeg để tải trực tiếp từ link m3u8/mpd và gộp thành mp4 (KHÔNG cần nén lại, cực nhanh)
        cmd = [
            "ffmpeg", "-y",
            "-headers", f"Cookie: {cookie_str}",
            "-i", url,
            "-c", "copy",  # Copy y nguyên luồng dữ liệu gốc, không nén lại
            str(output_path)
        ]
        
        # Bắt đầu chạy ffmpeg
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # Chờ ffmpeg tải xong (quá trình này phụ thuộc tốc độ mạng, thường rất nhanh)
        proc.wait(timeout=3600)  # Chờ tối đa 1 tiếng
        
        if proc.returncode == 0:
            size_mb = output_path.stat().st_size / (1024 * 1024)
            print(f"    [OK] Tải luồng thành công: {output_path.name} ({size_mb:.1f}MB)")
            return True
        else:
            stderr = proc.stderr.read().decode() if proc.stderr else ""
            print(f"    [ERROR] ffmpeg download failed: {stderr[-200:]}")
            return False
            
    except subprocess.TimeoutExpired:
        proc.kill()
        print("    [ERROR] ffmpeg download timeout.")
        return False
    except Exception as e:
        print(f"    [ERROR] Tải stream bị lỗi: {e}")
        return False

# ── Strategy 2: Screen recording with ffmpeg ─────────────────────────────────

def _parse_device_index(line: str) -> Optional[str]:
    """Extract device index from ffmpeg avfoundation line like '[...] [5] Device name'."""
    import re
    match = re.search(r'\[(\d+)\]', line)
    return match.group(1) if match else None


def _find_screen_device() -> Optional[str]:
    """Find the screen capture device index for ffmpeg avfoundation."""
    try:
        result = subprocess.run(
            ["ffmpeg", "-f", "avfoundation", "-list_devices", "true", "-i", ""],
            capture_output=True, text=True, timeout=10
        )
        for line in result.stderr.splitlines():
            if "Capture screen" in line:
                return _parse_device_index(line)
    except Exception:
        pass
    return None


def _find_audio_device() -> Optional[str]:
    """Find BlackHole or system audio device index for ffmpeg."""
    try:
        result = subprocess.run(
            ["ffmpeg", "-f", "avfoundation", "-list_devices", "true", "-i", ""],
            capture_output=True, text=True, timeout=10
        )
        in_audio = False
        for line in result.stderr.splitlines():
            if "AVFoundation audio devices" in line:
                in_audio = True
                continue
            # Chấp nhận BlackHole hoặc thiết bị Multi-Output (trường hợp user ko đổi tên)
            if in_audio and ("BlackHole" in line or "Soundflower" in line):
                return _parse_device_index(line)
            
            # Nếu tên thiết bị bị cắt xén hoặc là Multi-Output Device chung chung,
            # In ra log để xem nó đang có những thiết bị audio nào
            if in_audio and "]" in line:
                # Bỏ qua microphone của iPhone
                if "iPhone" not in line and "Microphone" not in line and "AirPods" not in line:
                    return _parse_device_index(line)
    except Exception:
        pass
    return None


def _get_video_duration(page: Page) -> Optional[float]:
    """Get video duration in seconds from the page's video element."""
    for frame in page.frames:
        try:
            duration = frame.evaluate("""
            () => {
                const v = document.querySelector('video');
                return v ? v.duration : null;
            }
            """)
            if duration and duration > 0:
                return float(duration)
        except Exception:
            pass
    return None


def record_screen(output_path: Path, duration: float, with_audio: bool = True) -> bool:
    """
    Record screen (and optionally audio) using ffmpeg.
    duration: recording duration in seconds.
    """
    if not shutil.which("ffmpeg"):
        print("    [ERROR] ffmpeg not found. Install: brew install ffmpeg")
        return False

    screen_dev = _find_screen_device()
    if not screen_dev:
        print("    [ERROR] No screen capture device found")
        return False

    audio_dev = _find_audio_device() if with_audio else None
    if with_audio and not audio_dev:
        print("    [WARN] No system audio device (BlackHole) found. Recording without audio.")
        print("    To enable audio: brew install blackhole-2ch")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if audio_dev:
        # CHIẾN LƯỢC TỐI THƯỢNG: TÁCH 2 LUỒNG VẬT LÝ VÀ CHỐNG CLIPPING
        cmd = [
            "ffmpeg", "-y",
            "-thread_queue_size", "2048",
            "-f", "avfoundation",
            "-framerate", "30",
            "-i", f"{screen_dev}:none",
            "-thread_queue_size", "2048",
            "-f", "avfoundation",
            "-i", f":{audio_dev}",
            "-c:v", "h264_videotoolbox",
            "-b:v", "4000k",
            "-pix_fmt", "yuv420p",
            "-c:a", "aac",
            "-b:a", "192k",
            "-ar", "48000",
            "-af", "volume=0.8,aresample=async=1",
            "-map", "0:v",
            "-map", "1:a"
        ]
    else:
        cmd = [
            "ffmpeg", "-y",
            "-thread_queue_size", "1024",
            "-f", "avfoundation",
            "-framerate", "30",
            "-i", f"{screen_dev}:none",
            "-c:v", "libx264",
            "-preset", "veryfast",
            "-crf", "26",
            "-pix_fmt", "yuv420p",
        ]
    cmd.append(str(output_path))

    print(f"    Recording screen for {duration:.0f}s → {output_path.name}")
    try:
        proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # Đợi trong đúng khoảng thời gian cần thiết thay vì dùng cờ -t của ffmpeg
        time.sleep(duration + 2)
        
        # Gửi phím "q" vào luồng nhập liệu để yêu cầu ffmpeg dừng lại một cách an toàn (lưu header mp4 đàng hoàng)
        try:
            if proc.stdin:
                proc.stdin.write(b"q\n")
                proc.stdin.flush()
        except Exception:
            pass

        proc.wait(timeout=15.0) # Đợi 15s để ffmpeg gói xong file
        
        if proc.returncode in (0, 255): # ffmpeg thoát khi bấm 'q' thường trả về 255 hoặc 0
            size_mb = output_path.stat().st_size / (1024 * 1024)
            print(f"    Saved: {output_path} ({size_mb:.1f}MB)")
            return True
        else:
            stderr = proc.stderr.read().decode() if proc.stderr else ""
            print(f"    [ERROR] ffmpeg failed (exit {proc.returncode}): {stderr[-200:]}")
            return False
    except subprocess.TimeoutExpired:
        proc.terminate()
        print("    [ERROR] ffmpeg timeout — force terminated")
        return False
    except Exception as e:
        print(f"    [ERROR] Screen recording failed: {e}")
        return False


# ── Main capture function ────────────────────────────────────────────────────

def capture_video_from_page(
    video_page: Page,
    output_path: Path,
    strategy: str = "auto",
    max_duration: Optional[float] = None,
) -> bool:
    """
    Capture a video from an already opened SharePoint video page.
    """
    output_path = output_path.with_suffix(".mp4")

    # Bỏ qua check skip khi đang test để ghi đè (luôn chạy lại)
    # if output_path.exists() and output_path.stat().st_size > 100_000:
    #    print(f"    [SKIP] Video already exists: {output_path}")
    #    return True

    from crawler import _activate_chrome
    _activate_chrome()
    video_page.bring_to_front()

    try:
        # Strategy 1: Network interception + download
        if strategy in ("auto", "intercept"):
            print("    [1/2] Intercepting video stream URL...")
            
            # Since page is already loaded, we just listen to network while forcing play
            captured_urls: list[str] = []
            
            def _on_response(response):
                u = response.url
                # Lắng nghe rộng hơn nữa: Bất cứ yêu cầu nào có chứa "videomanifest", "hls", "dash", "m3u8", "mpd", "mp4" hoặc định dạng mime type video
                ct = response.headers.get("content-type", "")
                if ("video/" in ct or
                    "application/dash+xml" in ct or
                    "application/vnd.apple.mpegurl" in ct or
                    ".mp4" in urlparse(u).path.lower() or 
                    "videomanifest" in u.lower() or 
                    "getvideostream" in u.lower() or
                    ".m3u8" in u.lower() or
                    ".mpd" in u.lower() or
                    "manifest" in u.lower()):
                    
                    # Bỏ qua các API tracking/telemetry của MS
                    if "telemetry" not in u.lower() and "events" not in u.lower() and "qos" not in u.lower():
                        captured_urls.append(u)
                    
            video_page.on("response", _on_response)
            
            try:
                # 1. Playwright mouse click
                video_page.mouse.click(video_page.viewport_size["width"] // 2, video_page.viewport_size["height"] // 2)
                time.sleep(1.0)
                
                # 2. Playwright locators
                for pb in video_page.locator('button, [role="button"]').filter(has_text="Play").all():
                    try: pb.click(timeout=1000)
                    except: pass
                
                # 3. JavaScript evaluation in all frames
                for frame in video_page.frames:
                    try:
                        frame.evaluate("""
                        () => {
                            const v = document.querySelector('video');
                            if (v && v.paused) v.play();
                        }
                        """)
                    except Exception:
                        pass
                        
                # 4. OS-level click using pyautogui (ultimate fallback for stubborn iframes)
                import pyautogui
                from crawler import _get_viewport_screen_region
                vp = _get_viewport_screen_region(video_page)
                if vp:
                    cx = vp["left"] + vp["width"] // 2
                    cy = vp["top"] + vp["height"] // 2
                    pyautogui.click(cx, cy)
                    time.sleep(0.5)
            except Exception:
                pass
            
            # Wait for video requests to appear
            deadline = time.time() + 30.0
            
            # Click nhấp nháy vài lần giữa màn hình để ép nó tải luồng mạng
            try:
                import pyautogui
                from crawler import _get_viewport_screen_region
                vp = _get_viewport_screen_region(video_page)
                if vp:
                    cx = vp["left"] + vp["width"] // 2
                    cy = vp["top"] + vp["height"] // 2
                    pyautogui.click(cx, cy)
            except:
                pass

            while time.time() < deadline and not captured_urls:
                time.sleep(1.0)
                
            try:
                video_page.remove_listener("response", _on_response)
            except Exception:
                pass

            # Lọc ưu tiên:
            # 1. Nếu có link trực tiếp MP4 (rất hiếm trên Stream)
            # 2. HLS (.m3u8) / DASH (.mpd) / manifest
            mp4_urls = [u for u in captured_urls if ".mp4" in urlparse(u).path.lower()]
            manifest_urls = [u for u in captured_urls if "manifest" in u.lower() or ".m3u8" in u.lower() or ".mpd" in u.lower()]
            video_content_urls = [u for u in captured_urls if "video/" in u.lower() or "application/" in u.lower()]
            
            video_url = None
            is_stream = False
            
            if mp4_urls:
                video_url = mp4_urls[0]
            elif manifest_urls:
                video_url = manifest_urls[0]
                is_stream = True
            elif video_content_urls:
                video_url = video_content_urls[0]
                if "getvideostream" in video_url.lower() or "videomanifest" in video_url.lower():
                    is_stream = True
            elif captured_urls:
                video_url = captured_urls[0]

            if video_url:
                print(f"    [INFO] Tìm thấy luồng Video URL: {video_url[:100]}...")
                if is_stream:
                    # Nếu là luồng Stream (HLS/DASH/m3u8/mpd), dùng ffmpeg để tải trực tiếp
                    if download_stream_ffmpeg(video_page, video_url, output_path):
                        return True
                else:
                    # Nếu là file MP4 tĩnh, thử dùng Playwright tải trực tiếp (vì có cookie auth)
                    if download_via_playwright(video_page, video_url, output_path):
                        return True
                    # Nếu thất bại, thử dùng urllib
                    if download_video(video_url, output_path):
                        return True
                print("    [WARN] Việc tải trực tiếp Video thất bại!")
            else:
                print("    [WARN] Không bắt được bất kỳ gói tin luồng Video URL nào")

            if strategy == "intercept":
                return False

        # Strategy 2: Screen recording
        if strategy in ("auto", "record"):
            print("    [2/2] Falling back to screen recording...")

            duration = _get_video_duration(video_page)
            if duration:
                print(f"    [INFO] Detected video duration from DOM: {duration} seconds")
            else:
                print("    [WARN] Could not detect video duration from DOM, using 10 minutes max")
                duration = 600.0

            if max_duration and duration > max_duration:
                print(f"    [TEST MODE] Limiting recording to {max_duration}s")
                duration = max_duration

            print(f"    Video duration: {duration:.0f}s ({duration / 60:.1f} min)")

            try:
                import pyautogui
            except Exception as e:
                print(f"    [ERROR] Không thể load pyautogui: {e}")

            # Đảm bảo tắt mọi âm thanh cảnh báo của mac để không lọt vào tiếng thu
            print("    [INFO] Bắt đầu kích hoạt video qua phím tắt...")
            
            try:
                from crawler import _get_viewport_screen_region
                vp = _get_viewport_screen_region(video_page)
                if vp:
                    cx = vp["left"] + vp["width"] // 2
                    cy = vp["top"] + vp["height"] // 2
                    print(f"    [INFO] Click chuột trái vào giữa màn hình ({cx}, {cy}) để focus...")
                    pyautogui.click(cx, cy)
                    time.sleep(1.0)
            except Exception as e:
                print(f"    [WARN] Lỗi khi cố gắng Click giữa màn hình: {e}")

            # Start video playback
            try:
                # 1. Bấm Option + K (Play/Pause)
                print("    [INFO] Bấm Option + K (Play/Pause video)...")
                pyautogui.hotkey("option", "k")
                time.sleep(2.0)
                
                # 2. Bật toàn màn hình player nội bộ bằng Option + Enter
                print("    [INFO] Bấm Option + Enter (Mở rộng toàn màn hình Player)...")
                pyautogui.hotkey("option", "enter")
                time.sleep(1.5)
                
                # Rê chuột ra rìa màn hình để ẩn UI của trình phát
                try:
                    from crawler import _get_viewport_screen_region
                    vp = _get_viewport_screen_region(video_page)
                    if vp:
                        pyautogui.moveTo(vp["left"] + 5, vp["top"] + 5)
                except:
                    pass
                
                # 3. JavaScript evaluation in all frames (Backup)
                for frame in video_page.frames:
                    try:
                        frame.evaluate("""
                        () => {
                            const v = document.querySelector('video');
                            if (v && v.paused) v.play();
                        }
                        """)
                    except Exception:
                        pass
            except Exception as e:
                print(f"    [WARN] Lỗi khi thao tác phím tắt video: {e}")
            time.sleep(1.0)

            # Cài đặt ffmpeg cuối cùng: Buffer 8MB, Codec ultrafast, AAC 128k, Async 1
            success = record_screen(output_path, duration, with_audio=True)

            # Exit fullscreen (Của Player)
            try:
                pyautogui.press("esc")
                time.sleep(0.5)
            except Exception as e:
                print(f"    [WARN] Lỗi khi thoát Fullscreen: {e}")

            return success

    finally:
        pass

    return False


def capture_video(
    page: Page,
    video_page_url: str,
    output_path: Path,
    strategy: str = "auto",
    max_duration: Optional[float] = None,
) -> bool:
    """
    Capture a video from SharePoint.

    strategy:
      "intercept" — network interception + direct download only
      "record"    — screen recording only
      "auto"      — try intercept first, fall back to record
    """
    output_path = output_path.with_suffix(".mp4")

    if output_path.exists() and output_path.stat().st_size > 100_000:
        print(f"    [SKIP] Video already exists: {output_path}")
        return True

    # Open video page in new tab
    new_page = page.context.new_page()
    new_page.bring_to_front()
    
    from crawler import _activate_chrome
    _activate_chrome()

    try:
        # Strategy 1: Network interception + download
        if strategy in ("auto", "intercept"):
            print("    [1/2] Intercepting video stream URL...")
            video_url = _intercept_video_url(new_page, video_page_url)

            if video_url:
                print(f"    Found video URL: {video_url[:100]}...")
                # Try browser-based download first (has auth cookies)
                if download_via_playwright(new_page, video_url, output_path):
                    return True
                # Fallback: direct download
                if download_video(video_url, output_path):
                    return True
                print("    [WARN] Direct download failed")
            else:
                print("    [WARN] Could not intercept video URL")

            if strategy == "intercept":
                return False

        # Strategy 2: Screen recording
        if strategy in ("auto", "record"):
            print("    [2/2] Falling back to screen recording...")

            # Navigate to video if not already there
            if new_page.url != video_page_url:
                new_page.goto(video_page_url, timeout=60000, wait_until="domcontentloaded")
                time.sleep(3.0)

            duration = _get_video_duration(new_page)
            if duration:
                print(f"    [INFO] Detected video duration from DOM: {duration} seconds")
            else:
                print("    [WARN] Could not detect video duration from DOM, using 10 minutes max")
                duration = 600.0

            if max_duration and duration > max_duration:
                print(f"    [TEST MODE] Limiting recording to {max_duration}s")
                duration = max_duration

            print(f"    Video duration: {duration:.0f}s ({duration / 60:.1f} min)")

            try:
                import pyautogui
            except Exception as e:
                print(f"    [ERROR] Không thể load pyautogui: {e}")

            # Fullscreen Chrome first so we click the true center of the video
            try:
                # Gửi tổ hợp phím Cmd+Ctrl+F để Safari/Chrome vào trạng thái Fullscreen
                pyautogui.hotkey("ctrl", "command", "f")
                time.sleep(2.0)
                
                # Focus trang bằng cách nhấp chuột TRÁI vào giữa màn hình
                from crawler import _get_viewport_screen_region
                vp = _get_viewport_screen_region(new_page)
                if vp:
                    cx = vp["left"] + vp["width"] // 2
                    cy = vp["top"] + vp["height"] // 2
                    print(f"    [INFO] Click chuột trái vào giữa màn hình ({cx}, {cy}) để focus và Play video...")
                    pyautogui.click(cx, cy)
                    time.sleep(1.0)
            except Exception as e:
                print(f"    [WARN] Lỗi khi cố gắng Fullscreen hoặc Click giữa màn hình: {e}")

            # Start video playback
            try:
                print("    [INFO] Bắt đầu kích hoạt video qua phím tắt...")
                
                # 1. Bật toàn màn hình player nội bộ bằng Option + Enter
                print("    [INFO] Bấm Option + Enter (Full screen player)...")
                pyautogui.hotkey("option", "enter")
                time.sleep(1.5)
                
                # Rê chuột ra rìa màn hình để ẩn UI của trình phát
                try:
                    from crawler import _get_viewport_screen_region
                    vp = _get_viewport_screen_region(new_page)
                    if vp:
                        pyautogui.moveTo(vp["left"] + 5, vp["top"] + 5)
                except:
                    pass
                
                # 3. Playwright mouse click as backup
                new_page.mouse.click(new_page.viewport_size["width"] // 2, new_page.viewport_size["height"] // 2)
                time.sleep(0.5)
                
                # 3. JavaScript evaluation in all frames
                for frame in new_page.frames:
                    try:
                        frame.evaluate("""
                        () => {
                            const v = document.querySelector('video');
                            if (v) { v.currentTime = 0; v.play(); }
                        }
                        """)
                    except Exception:
                        pass
            except Exception as e:
                print(f"    [WARN] Lỗi khi thao tác phím tắt video: {e}")
            time.sleep(1.0)

            success = record_screen(output_path, duration, with_audio=True)

            # Exit fullscreen (Cả player và Chrome)
            try:
                pyautogui.press("esc")
                time.sleep(0.5)
                pyautogui.hotkey("ctrl", "command", "f")
                time.sleep(0.5)
            except Exception as e:
                print(f"    [WARN] Lỗi khi thoát Fullscreen: {e}")

            return success

    finally:
        try:
            new_page.close()
        except Exception:
            pass

    return False
