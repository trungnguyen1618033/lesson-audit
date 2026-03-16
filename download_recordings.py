"""
download_recordings.py — Download MP4 files from SharePoint Recordings folder.

Cách dùng:
  uv run python download_recordings.py --url "URL_CỦA_THƯ_MỤC_RECORDINGS"

Script sẽ tạo thư mục `recordings/` và mỗi file MP4 sẽ được lưu vào `recordings/<Tên_file_không_đôi>/<Tên_file>.mp4`
"""

import sys
import time
import builtins
from pathlib import Path

# -- Override print to log to both console and recordings.log --
original_print = builtins.print

def log_print(*args, **kwargs):
    original_print(*args, **kwargs)
    
    # Don't log \r (carriage returns for progress bars)
    if kwargs.get("end", "") == "\r" or (args and isinstance(args[0], str) and args[0].startswith("\r")):
        return
        
    try:
        with open("recordings.log", "a", encoding="utf-8") as f:
            msg = " ".join(str(a) for a in args)
            f.write(msg + "\n")
    except Exception:
        pass

builtins.print = log_print

# -------------------------------------------------------------

from urllib.parse import urlparse, parse_qs
import click
from playwright.sync_api import sync_playwright

from crawler import (
    _activate_chrome,
    get_folder_items,
    get_server_relative_path,
    get_file_guid,
    _site_origin_base,
    slugify,
    wait_for_login,
    wait_for_sharepoint
)
from video_capture import capture_video, capture_video_from_page

RECORDINGS_DIR = Path("recordings")


def load_state() -> dict:
    import json
    state_file = Path("recordings_state.json")
    if state_file.exists():
        with open(state_file) as f:
            return json.load(f)
    return {"done": [], "failed": []}

def save_state(state: dict) -> None:
    import json
    with open("recordings_state.json", "w") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)

def log(msg: str) -> None:
    """Print to console and append to recordings.log"""
    print(msg)
    with open("recordings.log", "a", encoding="utf-8") as f:
        f.write(msg + "\n")

def process_recordings(page, url: str, target_file: str = None) -> None:
    log(f"\n============================================================")
    log(f"  Bắt đầu chạy lúc: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"  URL: {url}")
    if target_file:
        log(f"  Target: {target_file}")
    log(f"============================================================\n")
    
    log(f"Mở thư mục: {url}")
    page.goto(url, timeout=60000)
    wait_for_sharepoint(page)
    
    # Cuộn trang để load hết danh sách file (SharePoint lazy load)
    log("  Đang cuộn trang để lấy toàn bộ danh sách file...")
    for _ in range(5):
        page.mouse.wheel(0, 2000)
        time.sleep(1)
        
    items = get_folder_items(page)
    
    # Lọc file MP4
    mp4_files = [
        i for i in items
        if i["type"] == "file" and i["name"].lower().endswith(".mp4")
    ]
    
    if not mp4_files:
        log("  Không tìm thấy file MP4 nào trong thư mục này.")
        return
        
    log(f"  Tìm thấy {len(mp4_files)} file MP4.")
    
    state = load_state()
    RECORDINGS_DIR.mkdir(parents=True, exist_ok=True)
    
    for item in mp4_files:
        fname = item["name"]
        
        if target_file and fname != target_file:
            continue
            
        stem = Path(fname).stem
        
        # Cấu trúc: recordings/<Tên_file>/<Tên_file>.mp4
        safe_stem = slugify(stem)
        output_dir = RECORDINGS_DIR / safe_stem
        output_path = output_dir / f"{safe_stem}.mp4"
        
        queue_key = safe_stem
        
        # --- DEBUG MODE: Bỏ qua check file đã tồn tại để luôn luôn test ---
        # if queue_key in state["done"] and output_path.exists() and output_path.stat().st_size > 100_000:
        #    size_mb = output_path.stat().st_size / (1024 * 1024)
        #    log(f"  [SKIP] {fname} (đã download, {size_mb:.1f}MB)")
        #    continue
            
        log(f"\n  Đang xử lý: {fname}")
        
        try:
            from urllib.parse import quote
            srv_path = get_server_relative_path(url, fname)
            if srv_path:
                origin, base_path = _site_origin_base(url)
                # Create exact Stream player URL to bypass AllItems redirect
                encoded_path = quote(srv_path, safe='')
                # Using a dummy guid for referrerScenario just to make it valid
                video_url = f"{origin}{base_path}/_layouts/15/stream.aspx?id={encoded_path}&referrer=StreamWebApp.Web&referrerScenario=AddressBarCopied.view.3e4930d9-02e6-4055-9821-062dda219c4d"
            else:
                video_url = None
                log("    [ERROR] Không lấy được Server Relative Path.")
        except Exception as e:
            log(f"    [ERROR] Lỗi tạo URL: {e}")
            video_url = None
            
        if not video_url:
            log("    [ERROR] Bỏ qua file do không có URL.")
            continue

        try:
            log(f"    Mở tab mới với URL: {video_url[:80]}...")
            
            # --- PHƯƠNG PHÁP 3: CHẾ ĐỘ DOWNLOAD NATIVE (Tải trực tiếp qua trình duyệt) ---
            # Để giải quyết triệt để lỗi "nổ âm thanh" khi quay màn hình (do ffmpeg/Mac)
            # và lỗi "không bắt được gói tin" (do MS Stream mã hoá luồng HLS).
            # Chúng ta sẽ mở lại trang danh sách file (AllItems.aspx), 
            # click chọn file MP4 đó, và bấm nút "Download" trên thanh công cụ của SharePoint!
            
            log("    [INFO] Bắt đầu tiến trình tải file gốc qua nút Download của SharePoint...")
            
            # Mở một tab phụ để xử lý thao tác Download (sử dụng trang danh sách file gốc `url`)
            dl_page = page.context.new_page()
            dl_page.goto(url, wait_until="domcontentloaded", timeout=60000)
            time.sleep(3.0)
            
            # 1. Tìm cái thẻ <div> hoặc <span> có tên trùng khớp với fname trong danh sách
            # Sharepoint dùng grid ảo nên cần scroll
            target_locator = dl_page.locator(f"button[name='{fname}'], div[data-automationid='FieldRenderer-name']:has-text('{fname}')").first
            
            # Nếu không thấy ngay, phải scroll một chút
            if not target_locator.is_visible():
                dl_page.mouse.wheel(0, 500)
                time.sleep(1.0)
                if not target_locator.is_visible():
                     dl_page.mouse.wheel(0, 1000)
                     time.sleep(1.0)

            if target_locator.is_visible():
                # Bấm một phát để chọn (Select) cái file đó (Đánh dấu tick)
                target_locator.click()
                time.sleep(1.5)
                log("    [INFO] Đã tick chọn file trong danh sách.")
                
                # 2. Bấm nút Download trên thanh Ribbon
                download_btn = dl_page.locator("button[name='Download'], button[data-automationid='downloadCommand']").first
                if download_btn.is_visible():
                    log("    [INFO] Đã tìm thấy nút Download. Đang bắt đầu tải...")
                    
                    # Bắt sự kiện tải file của Playwright
                    with dl_page.expect_download(timeout=3600000) as download_info: # Chờ tối đa 1 tiếng để tải
                        download_btn.click()
                    
                    download = download_info.value
                    log(f"    [INFO] File đang được tải xuống từ máy chủ: {download.url[:80]}...")
                    
                    # Lưu file vào thư mục
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    download.save_as(str(output_path))
                    
                    size_mb = output_path.stat().st_size / (1024 * 1024)
                    log(f"    [OK] Tải file gốc thành công: {output_path.name} ({size_mb:.1f}MB)")
                    success = True
                else:
                    log("    [ERROR] Không tìm thấy nút Download trên thanh công cụ SharePoint.")
                    success = False
            else:
                log(f"    [ERROR] Không tìm thấy tên file '{fname}' trong danh sách để click chọn.")
                success = False

            dl_page.close()
            
            # --- KẾT THÚC PHƯƠNG PHÁP 3 ---
                
        except Exception as e:
            log(f"    [ERROR] Lỗi khi mở URL trong tab mới: {e}")
            continue
            
        if success:
            if queue_key not in state["done"]:
                state["done"].append(queue_key)
            save_state(state)
            log(f"    [OK] Hoàn tất: {output_path.name}")
        else:
            log(f"    [FAILED] Không thể capture: {fname}")


@click.command()
@click.option("--url", required=True, help="SharePoint URL trỏ tới thư mục Recordings.")
@click.option("--target", help="Tên file cụ thể muốn download (VD: abc.mp4).")
def main(url: str, target: str):
    """Download MP4 files từ SharePoint."""
    with sync_playwright() as p:
        print("Kết nối tới Chrome đang mở (cần khởi chạy với --remote-debugging-port=9222)...")
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        except Exception as e:
            print(f"Không kết nối được Chrome: {e}")
            sys.exit(1)
            
        context = browser.contexts[0]
        page = context.new_page()
        
        try:
            wait_for_login(page)
            process_recordings(page, url, target_file=target)
        except Exception as e:
            log(f"\n[CRITICAL ERROR] {e}")
        finally:
            page.close()
            log("\nHoàn tất quá trình.")

if __name__ == "__main__":
    main()
