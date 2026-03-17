import time
from pathlib import Path
from urllib.parse import quote

import click
from playwright.sync_api import sync_playwright

import video_capture
from crawler import get_folder_items, get_server_relative_path, _site_origin_base, wait_for_sharepoint, slugify

FOLDER_URL = (
    "https://hvtuphap.sharepoint.com/sites/LUTSHINHPQUCT8TP.HCM/Shared%20Documents/"
    "Forms/AllItems.aspx?id=%2Fsites%2FLUTSHINHPQUCT8TP%2EHCM%2FShared%20Documents"
    "%2FGeneral%2FRecordings&viewid=f40f232b%2Dbf1f%2D4cc1%2D940c%2D4282e3e23fe9"
    "&FolderCTID=0x01200033734350D7593049A68507234D499E6B"
)

TARGETS = [
    "Cuộc họp trong Chung-20250913_081001-Bản ghi cuộc họp.mp4",
    "Cuộc họp trong Chung-20250914_150233-Bản ghi cuộc họp.mp4",
]

OUTPUT_DIR = Path("captures/_test_7files")


def _has_video_element(video_page) -> bool:
    for frame in video_page.frames:
        try:
            if frame.evaluate("() => !!document.querySelector('video')"):
                return True
        except Exception:
            pass
    return False


def _open_by_click(page, fname):
    page.goto(FOLDER_URL, timeout=60000, wait_until="domcontentloaded")
    wait_for_sharepoint(page)
    time.sleep(2)

    unique_sub = fname.rsplit("-", 1)[0] if "-" in fname else fname.rsplit(".", 1)[0]
    for attempt in range(25):
        rect = page.evaluate("""(sub) => {
            const buttons = document.querySelectorAll('span[role="button"]');
            for (const btn of buttons) {
                const text = (btn.textContent || '').trim();
                if (text.includes(sub) && text.endsWith('.mp4')) {
                    const r = btn.getBoundingClientRect();
                    if (r.width > 0 && r.height > 0)
                        return { x: r.x + r.width / 2, y: r.y + r.height / 2, text: text };
                }
            }
            return null;
        }""", unique_sub)
        if rect:
            break
        page.mouse.wheel(0, 1500)
        time.sleep(0.8)

    if not rect:
        print(f"  Cannot find '{fname}' in folder DOM")
        return None

    new_pages = []
    page.context.on("page", lambda p: new_pages.append(p))
    page.mouse.click(rect["x"], rect["y"])

    deadline = time.time() + 15
    while not new_pages and time.time() < deadline:
        time.sleep(0.5)

    if not new_pages:
        print("  Click did not open a new tab")
        return None

    vp = new_pages[0]
    try:
        vp.wait_for_load_state("domcontentloaded", timeout=30000)
    except Exception:
        pass
    time.sleep(5)
    return vp


def open_video(page, fname):
    srv_path = get_server_relative_path(FOLDER_URL, fname)
    if not srv_path:
        return _open_by_click(page, fname)
    origin, base_path = _site_origin_base(FOLDER_URL)
    encoded_path = quote(srv_path, safe="")
    video_url = (
        f"{origin}{base_path}/_layouts/15/stream.aspx"
        f"?id={encoded_path}&referrer=StreamWebApp.Web"
        f"&referrerScenario=AddressBarCopied.view"
    )
    vp = page.context.new_page()
    vp.goto(video_url, timeout=60000, wait_until="domcontentloaded")
    time.sleep(5)

    if _has_video_element(vp):
        return vp

    print("  Constructed URL failed (no <video>), falling back to click from folder...")
    vp.close()
    return _open_by_click(page, fname)


def main():
    import sys
    sys.stdout.reconfigure(line_buffering=True)
    video_capture._CDP_DEBUG = False
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        print(f"Connected. {len(ctx.pages)} tab(s)\n")

        results = []
        for i, fname in enumerate(TARGETS):
            safe = slugify(Path(fname).stem) + ".mp4"
            out = OUTPUT_DIR / safe
            print(f"\n{'='*70}")
            print(f"[{i+1}/{len(TARGETS)}] {fname}")
            print(f"  Output: {out}")

            vp = None
            t0 = time.time()
            try:
                vp = open_video(page, fname)
                if not vp:
                    raise RuntimeError("Cannot build video URL")

                ok = video_capture.capture_video_from_page(
                    video_page=vp, output_path=out,
                    strategy="auto", max_duration=30,
                )
                elapsed = time.time() - t0
                if ok and out.exists() and out.stat().st_size > 100_000:
                    sz = out.stat().st_size / (1024*1024)
                    print(f"  OK ({sz:.1f}MB, {elapsed:.0f}s)")
                    results.append((fname, "OK", f"{sz:.1f}MB", f"{elapsed:.0f}s"))
                else:
                    print(f"  FAILED ({elapsed:.0f}s)")
                    results.append((fname, "FAILED", "-", f"{elapsed:.0f}s"))
            except Exception as e:
                elapsed = time.time() - t0
                print(f"  ERROR: {e} ({elapsed:.0f}s)")
                results.append((fname, f"ERROR: {e}"[:60], "-", f"{elapsed:.0f}s"))
            finally:
                if vp:
                    try: vp.close()
                    except: pass

        print(f"\n{'='*70}")
        print("SUMMARY:")
        for fname, status, size, elapsed in results:
            short = fname[:55] + "..." if len(fname) > 55 else fname
            print(f"  {status:8s} {size:>8s} {elapsed:>6s}  {short}")


if __name__ == "__main__":
    main()
