"""
video_capture.py -- Capture video (MP4) files from SharePoint.

Three strategies (auto = 1 then 2):
  1. Network interception: Intercept video stream URL via Playwright, download directly.
  2. CDP packet capture: Intercept actual response bodies via Chrome DevTools Protocol
     while the video plays, then reassemble segments with ffmpeg. Original quality,
     no re-encoding.
  3. Screen recording: Record screen + system audio with ffmpeg.
     Only used when explicitly requested (strategy="record").

macOS audio note: System audio capture requires BlackHole virtual audio device.
  Install: brew install blackhole-2ch
  Then create a Multi-Output Device in Audio MIDI Setup.
"""

import base64
import shutil
import subprocess
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
        file_size = output_path.stat().st_size
        if file_size < 100_000:
            print(f"    [WARN] File too small ({file_size} bytes), not a valid video")
            output_path.unlink()
            return False
        size_mb = file_size / (1024 * 1024)
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
            data = resp.body()
            if len(data) < 100_000:
                print(f"    [WARN] Response too small ({len(data)} bytes), likely a stream segment, not full file")
                return False
            output_path.write_bytes(data)
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

# ── Strategy 3: CDP packet capture ────────────────────────────────────────────

_MSE_HOOK_JS = """
(() => {
    if (window.__mseCapture) return;
    window.__mseCapture = { tracks: [] };
    const trackMap = new WeakMap();
    const origAddSB = MediaSource.prototype.addSourceBuffer;
    MediaSource.prototype.addSourceBuffer = function(mimeType) {
        const sb = origAddSB.call(this, mimeType);
        const track = { mimeType: mimeType, chunks: [], totalBytes: 0 };
        window.__mseCapture.tracks.push(track);
        trackMap.set(sb, track);
        return sb;
    };
    const origAppend = SourceBuffer.prototype.appendBuffer;
    SourceBuffer.prototype.appendBuffer = function(data) {
        const track = trackMap.get(this);
        if (track) {
            const bytes = (data instanceof ArrayBuffer) ? new Uint8Array(data)
                : (data instanceof Uint8Array) ? data
                : new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
            track.chunks.push(bytes.slice());
            track.totalBytes += bytes.byteLength;
        }
        return origAppend.call(this, data);
    };
})();
"""


_CDP_DEBUG = True


def _cdp_log(msg: str) -> None:
    if _CDP_DEBUG:
        print(f"    [CDP:DBG] {msg}")


def _is_video_segment(url: str, content_type: str) -> bool:
    url_lower = url.lower()
    ct_lower = content_type.lower()

    skip = ("telemetry", "events", "qos", "logging", "beacon", "analytics", "clienttelemetry")
    if any(s in url_lower for s in skip):
        return False

    if any(t in ct_lower for t in ("video/", "audio/mp4", "audio/webm")):
        return True

    video_url_markers = (
        "qualitylevels(", "fragments(", "getvideostream",
        ".m4s", "/segment", "/chunk", "videoplayback",
        "range/", "video=", "audio=",
    )
    if any(m in url_lower for m in video_url_markers):
        if "application/octet-stream" in ct_lower or "binary" in ct_lower or not ct_lower:
            return True

    if ".mp4" in urlparse(url).path.lower() and "manifest" not in url_lower:
        return True

    return False


def _inject_mse_hook(page: Page) -> None:
    injected = 0
    for frame in page.frames:
        try:
            frame.evaluate(_MSE_HOOK_JS)
            injected += 1
        except Exception:
            pass
    _cdp_log(f"MSE hook injected into {injected}/{len(page.frames)} frame(s)")


def _collect_mse_tracks(page: Page, tmp_dir: Path) -> list[tuple[Path, str]]:
    """Collect MSE-captured tracks from all frames. Returns [(file_path, mime_type), ...]."""
    results: list[tuple[Path, str]] = []
    BATCH_BYTES = 5 * 1024 * 1024

    for frame in page.frames:
        try:
            tracks_info = frame.evaluate("""
            () => {
                if (!window.__mseCapture) return [];
                return window.__mseCapture.tracks.map((t, i) => ({
                    index: i, mimeType: t.mimeType,
                    numChunks: t.chunks.length, totalBytes: t.totalBytes
                }));
            }
            """)
        except Exception:
            continue

        if not tracks_info:
            continue

        for track in tracks_info:
            if track["numChunks"] == 0:
                continue

            mime = track["mimeType"]
            suffix = ".video.bin" if "video" in mime else ".audio.bin"
            track_path = tmp_dir / f"mse_track_{track['index']}{suffix}"

            with open(track_path, "wb") as f:
                chunk_idx = 0
                while chunk_idx < track["numChunks"]:
                    result = frame.evaluate(f"""
                    () => {{
                        const track = window.__mseCapture.tracks[{track['index']}];
                        let batch = [];
                        let batchSize = 0;
                        let idx = {chunk_idx};
                        while (idx < track.chunks.length && batchSize < {BATCH_BYTES}) {{
                            batch.push(track.chunks[idx]);
                            batchSize += track.chunks[idx].byteLength;
                            idx++;
                        }}
                        const combined = new Uint8Array(batchSize);
                        let offset = 0;
                        for (const c of batch) {{
                            combined.set(c, offset);
                            offset += c.byteLength;
                        }}
                        const SZ = 32768;
                        let binary = '';
                        for (let i = 0; i < combined.length; i += SZ) {{
                            binary += String.fromCharCode.apply(
                                null, combined.subarray(i, Math.min(i + SZ, combined.length)));
                        }}
                        return {{ data: btoa(binary), consumed: batch.length }};
                    }}
                    """)
                    raw = base64.b64decode(result["data"])
                    f.write(raw)
                    chunk_idx += result["consumed"]

            if track_path.stat().st_size > 0:
                results.append((track_path, mime))

        if results:
            break

    return results


def _start_playback(page: Page) -> None:
    try:
        page.mouse.click(page.viewport_size["width"] // 2, page.viewport_size["height"] // 2)
        time.sleep(1.0)
    except Exception:
        pass

    try:
        for pb in page.locator('button, [role="button"]').filter(has_text="Play").all():
            try:
                pb.click(timeout=1000)
            except Exception:
                pass
    except Exception:
        pass

    for frame in page.frames:
        try:
            frame.evaluate("() => { const v = document.querySelector('video'); if (v && v.paused) { v.currentTime = 0; v.play(); } }")
        except Exception:
            pass

    try:
        import pyautogui
        from crawler import _get_viewport_screen_region
        vp = _get_viewport_screen_region(page)
        if vp:
            cx = vp["left"] + vp["width"] // 2
            cy = vp["top"] + vp["height"] // 2
            pyautogui.click(cx, cy)
    except Exception:
        pass

    time.sleep(1.0)


def _reassemble_segments(
    segment_files: list[dict],
    mse_tracks: list[tuple[Path, str]],
    output_path: Path,
) -> bool:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not shutil.which("ffmpeg"):
        print("    [ERROR] ffmpeg not found. Install: brew install ffmpeg")
        return False

    if mse_tracks:
        print(f"    [CDP] Reassembling from {len(mse_tracks)} MSE track(s)...")
        cmd = ["ffmpeg", "-y"]
        for track_path, _ in mse_tracks:
            cmd.extend(["-i", str(track_path)])
        cmd.extend(["-c", "copy"])
        for i in range(len(mse_tracks)):
            cmd.extend(["-map", str(i)])
        cmd.append(str(output_path))

        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if proc.returncode == 0 and output_path.exists() and output_path.stat().st_size > 10_000:
                size_mb = output_path.stat().st_size / (1024 * 1024)
                print(f"    [CDP] Reassembled (MSE): {output_path.name} ({size_mb:.1f}MB)")
                return True
            _cdp_log(f"MSE ffmpeg exit={proc.returncode} stderr={proc.stderr[-500:]}")
        except Exception as e:
            _cdp_log(f"MSE reassembly exception: {e}")

    if not segment_files:
        print("    [ERROR] No segments captured")
        return False

    sorted_segs = sorted(segment_files, key=lambda s: s["seq"])
    print(f"    [CDP] Reassembling from {len(sorted_segs)} CDP segment(s)...")

    video_segs: list[dict] = []
    audio_segs: list[dict] = []
    other_segs: list[dict] = []
    for seg in sorted_segs:
        ct = seg.get("content_type", "").lower()
        url = seg.get("url", "").lower()
        if "audio" in ct or "audio=" in url:
            audio_segs.append(seg)
        elif "video" in ct or "video=" in url:
            video_segs.append(seg)
        else:
            other_segs.append(seg)

    if not video_segs and not audio_segs:
        video_segs = other_segs
        other_segs = []
    elif other_segs:
        video_segs.extend(other_segs)
        video_segs.sort(key=lambda s: s["seq"])

    tmp_dir = sorted_segs[0]["path"].parent

    def _concat(segs: list[dict], name: str) -> Path:
        out = tmp_dir / name
        with open(out, "wb") as f:
            for s in segs:
                f.write(s["path"].read_bytes())
        return out

    inputs = []
    if video_segs:
        inputs.append(_concat(video_segs, "combined_video.bin"))
    if audio_segs:
        inputs.append(_concat(audio_segs, "combined_audio.bin"))

    _cdp_log(f"CDP reassembly: {len(video_segs)} video seg(s), {len(audio_segs)} audio seg(s)")
    for inp in inputs:
        _cdp_log(f"  Input: {inp.name} ({inp.stat().st_size / (1024*1024):.1f}MB)")

    cmd = ["ffmpeg", "-y"]
    for inp in inputs:
        cmd.extend(["-i", str(inp)])
    cmd.extend(["-c", "copy"])
    for i in range(len(inputs)):
        cmd.extend(["-map", str(i)])
    cmd.append(str(output_path))

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if proc.returncode == 0 and output_path.exists() and output_path.stat().st_size > 10_000:
            size_mb = output_path.stat().st_size / (1024 * 1024)
            print(f"    [CDP] Reassembled (CDP): {output_path.name} ({size_mb:.1f}MB)")
            return True
        _cdp_log(f"ffmpeg exit={proc.returncode} stderr={proc.stderr[-500:]}")
    except Exception as e:
        _cdp_log(f"CDP reassembly exception: {e}")

    if video_segs and not audio_segs:
        try:
            concat_str = "|".join(str(s["path"]) for s in sorted_segs)
            cmd2 = ["ffmpeg", "-y", "-i", f"concat:{concat_str}", "-c", "copy", str(output_path)]
            proc2 = subprocess.run(cmd2, capture_output=True, text=True, timeout=300)
            if proc2.returncode == 0 and output_path.exists() and output_path.stat().st_size > 10_000:
                size_mb = output_path.stat().st_size / (1024 * 1024)
                print(f"    [CDP] Reassembled (concat): {output_path.name} ({size_mb:.1f}MB)")
                return True
            _cdp_log(f"concat ffmpeg exit={proc2.returncode} stderr={proc2.stderr[-500:]}")
        except Exception as e:
            _cdp_log(f"Concat reassembly exception: {e}")

    print("    [ERROR] All reassembly methods failed")
    return False


def _mute_and_speed(page: Page, rate: float = 4.0) -> float:
    applied_rate = 1.0
    for frame in page.frames:
        try:
            applied_rate = frame.evaluate("""(rate) => {
                const v = document.querySelector('video');
                if (!v) return 1.0;
                v.muted = true;
                v.playbackRate = rate;
                return v.playbackRate;
            }""", rate)
            if applied_rate > 1.0:
                break
        except Exception:
            pass
    return applied_rate


def capture_video_via_cdp(
    page: Page,
    output_path: Path,
    max_duration: Optional[float] = None,
    playback_rate: float = 16.0,
) -> bool:
    output_path = output_path.with_suffix(".mp4")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_dir = Path(tempfile.mkdtemp(prefix="video_cdp_"))
    cdp = None

    try:
        print("    [CDP] Setting up network packet capture...")
        cdp = page.context.new_cdp_session(page)
        cdp.send("Network.enable", {
            "maxResourceBufferSize": 100 * 1024 * 1024,
            "maxTotalBufferSize": 500 * 1024 * 1024,
        })

        video_request_meta: dict[str, dict] = {}
        pending_ids: list[str] = []
        segment_files: list[dict] = []
        seq_counter = [0]
        all_response_count = [0]
        skipped_urls: list[str] = []

        def _on_response_received(params):
            all_response_count[0] += 1
            url = params.get("response", {}).get("url", "")
            headers = params.get("response", {}).get("headers", {})
            ct = headers.get("content-type", headers.get("Content-Type", ""))
            cr = headers.get("content-range", headers.get("Content-Range", ""))
            status = params.get("response", {}).get("status", 0)

            is_match = _is_video_segment(url, ct)
            if is_match:
                rid = params["requestId"]
                video_request_meta[rid] = {
                    "url": url,
                    "content_type": ct,
                    "content_range": cr,
                    "seq": seq_counter[0],
                    "status": status,
                }
                seq_counter[0] += 1
                url_short = url[:120] + ("..." if len(url) > 120 else "")
                _cdp_log(f"MATCH #{seq_counter[0]} [{status}] ct={ct!r} range={cr!r} url={url_short}")
            else:
                if ct and ("video" in ct.lower() or "audio" in ct.lower() or "octet" in ct.lower()):
                    url_short = url[:120] + ("..." if len(url) > 120 else "")
                    skipped_urls.append(url_short)
                    _cdp_log(f"SKIP (filter miss?) [{status}] ct={ct!r} url={url_short}")

        def _on_loading_finished(params):
            rid = params["requestId"]
            if rid in video_request_meta:
                pending_ids.append(rid)

        cdp.on("Network.responseReceived", _on_response_received)
        cdp.on("Network.loadingFinished", _on_loading_finished)

        body_fail_count = [0]

        def _drain_pending():
            while pending_ids:
                rid = pending_ids.pop(0)
                meta = video_request_meta.get(rid)
                if not meta:
                    continue
                try:
                    body_resp = cdp.send("Network.getResponseBody", {"requestId": rid})
                    is_b64 = body_resp.get("base64Encoded", False)
                    raw = base64.b64decode(body_resp["body"]) if is_b64 else body_resp["body"].encode("latin-1")
                    seg_path = tmp_dir / f"seg_{meta['seq']:05d}.bin"
                    seg_path.write_bytes(raw)
                    segment_files.append({
                        "path": seg_path,
                        "seq": meta["seq"],
                        "url": meta["url"],
                        "content_type": meta["content_type"],
                        "content_range": meta["content_range"],
                        "size": len(raw),
                    })
                    total_mb = sum(s["size"] for s in segment_files) / (1024 * 1024)
                    print(f"\r    [CDP] Captured {len(segment_files)} segments ({total_mb:.1f}MB)", end="", flush=True)
                except Exception as e:
                    body_fail_count[0] += 1
                    url_short = meta["url"][:80]
                    _cdp_log(f"getResponseBody FAILED #{body_fail_count[0]} seq={meta['seq']} err={e} url={url_short}")

        page.add_init_script(_MSE_HOOK_JS)
        _cdp_log("MSE hook registered via add_init_script (will run on reload)")

        print("    [CDP] Reloading page to capture init segments...")
        page.reload(wait_until="domcontentloaded", timeout=60000)
        time.sleep(5)
        _drain_pending()
        _cdp_log(f"After reload: {len(segment_files)} init segments captured")

        _inject_mse_hook(page)

        print("    [CDP] Starting video playback...")
        _start_playback(page)

        actual_rate = _mute_and_speed(page, playback_rate)
        if actual_rate > 1.0:
            print(f"    [CDP] Muted + playback rate: {actual_rate}x")
        else:
            print("    [CDP] Muted (playback rate unchanged)")

        print("    [CDP] Capturing video packets...")
        duration = _get_video_duration(page) or 0
        if max_duration and (duration <= 0 or duration > max_duration):
            duration = max_duration
        if duration <= 0:
            duration = 600.0
        print(f"    [CDP] Video duration: {duration:.0f}s ({duration / 60:.1f} min)")

        wall_timeout = duration / actual_rate + 30
        start_time = time.time()
        last_capture_time = time.time()
        last_count = 0
        last_state_log = 0.0
        IDLE_TIMEOUT = 15.0

        while True:
            page.wait_for_timeout(2000)
            _drain_pending()

            elapsed = time.time() - start_time

            if elapsed - last_state_log >= 10.0:
                last_state_log = elapsed
                video_state = None
                for frame in page.frames:
                    try:
                        video_state = frame.evaluate("""
                        () => {
                            const v = document.querySelector('video');
                            if (!v) return null;
                            return {
                                paused: v.paused, ended: v.ended,
                                currentTime: Math.round(v.currentTime),
                                duration: Math.round(v.duration || 0),
                                readyState: v.readyState, networkState: v.networkState,
                                buffered: v.buffered.length > 0
                                    ? Math.round(v.buffered.end(v.buffered.length - 1)) : 0
                            };
                        }
                        """)
                        if video_state:
                            break
                    except Exception:
                        pass
                if video_state:
                    vs = video_state
                    state_str = "PAUSED" if vs["paused"] else ("ENDED" if vs["ended"] else "PLAYING")
                    _cdp_log(
                        f"t={elapsed:.0f}s video={state_str} pos={vs['currentTime']}/{vs['duration']}s "
                        f"buffered={vs['buffered']}s ready={vs['readyState']} net={vs['networkState']} "
                        f"segs={len(segment_files)} reqs={all_response_count[0]} fails={body_fail_count[0]}"
                    )
                else:
                    _cdp_log(f"t={elapsed:.0f}s <video> element NOT FOUND in any frame")

            if elapsed > wall_timeout:
                print(f"\n    [CDP] Duration timeout ({duration:.0f}s / {actual_rate}x + 30s = {wall_timeout:.0f}s)")
                break

            ended = False
            for frame in page.frames:
                try:
                    ended = frame.evaluate("() => { const v = document.querySelector('video'); return v ? v.ended : false; }")
                    if ended:
                        break
                except Exception:
                    pass

            if ended:
                print(f"\n    [CDP] Video ended at {elapsed:.0f}s")
                page.wait_for_timeout(3000)
                _drain_pending()
                break

            cur_count = len(segment_files)
            if cur_count > last_count:
                last_count = cur_count
                last_capture_time = time.time()
            elif cur_count > 0 and (time.time() - last_capture_time) > IDLE_TIMEOUT:
                print(f"\n    [CDP] No new segments for {IDLE_TIMEOUT:.0f}s, assuming complete")
                break

        try:
            cdp.send("Network.disable")
        except Exception:
            pass

        print()
        total_mb = sum(s["size"] for s in segment_files) / (1024 * 1024) if segment_files else 0
        print(f"    [CDP] Capture complete: {len(segment_files)} segments, {total_mb:.1f}MB")
        _cdp_log(f"Summary: total_responses={all_response_count[0]} matched={seq_counter[0]} "
                 f"saved={len(segment_files)} body_fails={body_fail_count[0]}")
        if skipped_urls:
            _cdp_log(f"Skipped {len(skipped_urls)} responses with media-like content-type:")
            for u in skipped_urls[:5]:
                _cdp_log(f"  {u}")

        if segment_files:
            ct_counts: dict[str, int] = {}
            for s in segment_files:
                ct = s["content_type"] or "(empty)"
                ct_counts[ct] = ct_counts.get(ct, 0) + 1
            _cdp_log(f"Segment content-types: {ct_counts}")

        mse_tracks = _collect_mse_tracks(page, tmp_dir)
        if mse_tracks:
            mse_total = sum(p.stat().st_size for p, _ in mse_tracks) / (1024 * 1024)
            print(f"    [CDP] MSE backup: {len(mse_tracks)} track(s), {mse_total:.1f}MB")
            for tp, mime in mse_tracks:
                _cdp_log(f"  MSE track: {tp.name} mime={mime} size={tp.stat().st_size / (1024*1024):.1f}MB")
        else:
            _cdp_log("MSE hook captured 0 tracks")

        if not segment_files and not mse_tracks:
            print("    [ERROR] No video data captured via CDP or MSE")
            _cdp_log(f"DIAGNOSTIC: {all_response_count[0]} total responses seen, 0 matched video filter")
            return False

        return _reassemble_segments(segment_files, mse_tracks, output_path)

    finally:
        try:
            if cdp:
                cdp.detach()
        except Exception:
            pass
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


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

    strategy:
      "auto"        -- try intercept, then cdp_capture
      "intercept"   -- network interception + direct download only
      "cdp_capture" -- CDP packet capture only
      "record"      -- screen recording only (use as last resort)
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

            _start_playback(video_page)

            deadline = time.time() + 30.0
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

        # Strategy 3: CDP packet capture
        if strategy in ("auto", "cdp_capture"):
            print("    [2/2] CDP packet capture...")
            if capture_video_via_cdp(video_page, output_path, max_duration=max_duration):
                return True
            print("    [WARN] CDP packet capture failed")
            if strategy == "cdp_capture":
                return False

        # Strategy 2: Screen recording (only when explicitly requested)
        if strategy == "record":
            print("    [RECORD] Screen recording...")

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

            try:
                print("    [INFO] Bấm Option + K (Play/Pause video)...")
                pyautogui.hotkey("option", "k")
                time.sleep(2.0)
                
                print("    [INFO] Bấm Option + Enter (Mở rộng toàn màn hình Player)...")
                pyautogui.hotkey("option", "enter")
                time.sleep(1.5)
                
                try:
                    from crawler import _get_viewport_screen_region
                    vp = _get_viewport_screen_region(video_page)
                    if vp:
                        pyautogui.moveTo(vp["left"] + 5, vp["top"] + 5)
                except:
                    pass
                
                for frame in video_page.frames:
                    try:
                        frame.evaluate("() => { const v = document.querySelector('video'); if (v && v.paused) v.play(); }")
                    except Exception:
                        pass
            except Exception as e:
                print(f"    [WARN] Lỗi khi thao tác phím tắt video: {e}")
            time.sleep(1.0)

            success = record_screen(output_path, duration, with_audio=True)

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
      "auto"        -- try intercept, then cdp_capture
      "intercept"   -- network interception + direct download only
      "cdp_capture" -- CDP packet capture (intercept response bodies while playing)
      "record"      -- screen recording only (use as last resort)
    """
    output_path = output_path.with_suffix(".mp4")

    if output_path.exists() and output_path.stat().st_size > 100_000:
        print(f"    [SKIP] Video already exists: {output_path}")
        return True

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
                if download_via_playwright(new_page, video_url, output_path):
                    return True
                if download_video(video_url, output_path):
                    return True
                print("    [WARN] Direct download failed")
            else:
                print("    [WARN] Could not intercept video URL")

            if strategy == "intercept":
                return False

        # Strategy 3: CDP packet capture
        if strategy in ("auto", "cdp_capture"):
            print("    [2/2] CDP packet capture...")
            if new_page.url != video_page_url:
                new_page.goto(video_page_url, timeout=60000, wait_until="domcontentloaded")
                time.sleep(3.0)
            if capture_video_via_cdp(new_page, output_path, max_duration=max_duration):
                return True
            print("    [WARN] CDP packet capture failed")
            if strategy == "cdp_capture":
                return False

        # Strategy 2: Screen recording (only when explicitly requested)
        if strategy == "record":
            print("    [RECORD] Screen recording...")

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

            try:
                pyautogui.hotkey("ctrl", "command", "f")
                time.sleep(2.0)
                
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

            try:
                print("    [INFO] Bắt đầu kích hoạt video qua phím tắt...")
                
                print("    [INFO] Bấm Option + Enter (Full screen player)...")
                pyautogui.hotkey("option", "enter")
                time.sleep(1.5)
                
                try:
                    from crawler import _get_viewport_screen_region
                    vp = _get_viewport_screen_region(new_page)
                    if vp:
                        pyautogui.moveTo(vp["left"] + 5, vp["top"] + 5)
                except:
                    pass
                
                new_page.mouse.click(new_page.viewport_size["width"] // 2, new_page.viewport_size["height"] // 2)
                time.sleep(0.5)
                
                for frame in new_page.frames:
                    try:
                        frame.evaluate("() => { const v = document.querySelector('video'); if (v) { v.currentTime = 0; v.play(); } }")
                    except Exception:
                        pass
            except Exception as e:
                print(f"    [WARN] Lỗi khi thao tác phím tắt video: {e}")
            time.sleep(1.0)

            success = record_screen(output_path, duration, with_audio=True)

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
