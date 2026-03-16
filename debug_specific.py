import time
from playwright.sync_api import sync_playwright
from crawler import (
    get_folder_items, sharepoint_subfolder_url, wait_for_sharepoint,
    open_pptx_and_present, CDP_URL, load_queue, _CHECK_BTNS_JS
)

# Cấu hình file lỗi user đang gặp
TARGET_FILE = "Tong quan, nhan thuc hien vu viec.pptx"
SUBJECT_NAME = "1. NLS"
SESSION_HINT = "01.03" 

def main():
    state = load_queue()
    root_url = state.get("root_url")
    if not root_url:
        print("Error: No root_url in queue_state.json. Run crawler normally first.")
        return

    print(f"Looking for '{TARGET_FILE}' in subject '{SUBJECT_NAME}'...")
    print(f"Connecting to Chrome at {CDP_URL}...")
    
    with sync_playwright() as pw:
        try:
            browser = pw.chromium.connect_over_cdp(CDP_URL)
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = context.pages[0] if context.pages else context.new_page()
        except Exception as e:
            print(f"Error connecting to Chrome: {e}")
            return

        # 1. Go to Subject
        subject_url = sharepoint_subfolder_url(root_url, SUBJECT_NAME)
        print(f"Navigating to subject: {subject_url}")
        page.goto(subject_url)
        wait_for_sharepoint(page)
        
        # 2. List sessions
        items = get_folder_items(page)
        sessions = [i for i in items if i["type"] == "folder"]
        
        target_session_url = None
        target_item = None
        
        print(f"Scanning sessions matching '{SESSION_HINT}'...")
        for sess in sessions:
            if SESSION_HINT and SESSION_HINT not in sess["name"]:
                continue
                
            sess_url = sharepoint_subfolder_url(subject_url, sess["name"])
            print(f"Checking session: {sess['name']} ...")
            page.goto(sess_url)
            wait_for_sharepoint(page)
            
            files = get_folder_items(page)
            for f in files:
                if f["name"] == TARGET_FILE:
                    target_item = f
                    target_session_url = sess_url
                    print(f"FOUND FILE in {sess['name']}!")
                    break
            if target_item:
                break
                
        if not target_item:
            print(f"File '{TARGET_FILE}' not found.")
            return

        # 3. Test Open & Present with DEBUGGING
        print(f"\n--- Testing Open & Present for: {TARGET_FILE} ---")
        
        # Manually open to inspect buttons
        from crawler import get_server_relative_path, get_file_guid, build_office_url
        
        srv_path = get_server_relative_path(target_session_url, TARGET_FILE)
        guid = get_file_guid(page, srv_path)
        target_url = build_office_url(guid, TARGET_FILE, target_session_url, action="view")
        
        print(f"Opening URL: {target_url}")
        new_page = page.context.new_page()
        new_page.goto(target_url)
        new_page.bring_to_front()
        
        print("Waiting 10s for page load...")
        time.sleep(10)
        
        print("\n--- BUTTON INSPECTION ---")
        found_any = False
        for i, frame in enumerate(new_page.frames):
            try:
                # Dump all buttons with 'Present' text
                btns = frame.locator('button, [role="button"], [role="menuitem"], [role="tab"]').filter(
                    has_text=re.compile(r'Trình bày|Present|Start Slide Show|From Beginning|Từ đầu', re.IGNORECASE)
                ).all()
                
                if btns:
                    print(f"Frame {i} ({frame.url[:40]}...): Found {len(btns)} candidate buttons.")
                    found_any = True
                    for j, btn in enumerate(btns):
                        txt = btn.text_content() or ""
                        aria = btn.get_attribute("aria-label") or ""
                        vis = btn.is_visible()
                        en = btn.is_enabled()
                        box = btn.bounding_box()
                        print(f"  [{j}] Text: '{txt.strip()}' | Aria: '{aria}' | Visible: {vis} | Enabled: {en} | Box: {box}")
                        
                        # Highlighting
                        if box:
                            frame.evaluate(f"""
                                const el = document.querySelectorAll('button, [role="button"], [role="menuitem"], [role="tab"]')
                                           [{j}]; // This is approximate, locator.all() returns elements handle
                            """)
                            # Better highlight via PyAutoGUI not possible precisely without screen coords
                            # But we can try to click the first visible enabled one
            except Exception as e:
                pass
                
        if not found_any:
            print("No 'Present' buttons found via Locator.")
            
        print("\nAttempting standard open_pptx_and_present logic...")
        # Close and use standard function to test fallback
        new_page.close()
        new_page = open_pptx_and_present(page, TARGET_FILE, session_url=target_session_url)

        if new_page:
            print("\n✅ SUCCESS: Slideshow started.")
            print("Check Chrome window now.")
            input("Press Enter to finish...")
            try: new_page.close()
            except: pass
        else:
            print("\n❌ FAILED.")

import re

if __name__ == "__main__":
    main()
