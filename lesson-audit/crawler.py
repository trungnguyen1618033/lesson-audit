"""
crawler.py - SharePoint Auto Crawler

Tự động duyệt cấu trúc thư mục SharePoint, tìm file PPTX trong mỗi buổi học,
mở từng file bằng PowerPoint Online, rồi gọi capture tool để chụp và xuất PPTX.

Cấu trúc SharePoint:
  2. Tai lieu bai giang/
  ├── 1.NLS/
  │   ├── 1.Sang Thu Bay 22.02/   ← buổi học
  │   │   ├── file.pptx
  │   │   └── file.docx
  │   └── ...
  └── 2.KNC8/ ...

Output:
  captures/1.NLS/1.Sang_Thu_Bay_22.02/TenFile/slide_001.png ...

Cách dùng:
  # Bước 1: Mở Chrome với remote debugging (PHẢI dùng --user-data-dir)
  /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \\
    --remote-debugging-port=9222 \\
    --user-data-dir=/tmp/chrome-crawl

  # Bước 2: Đăng nhập SharePoint trong Chrome vừa mở, sau đó chạy crawler
  uv run python crawler.py --url "URL_TRANG_TAI_LIEU_BAI_GIANG"

  # Resume nếu bị ngắt
  uv run python crawler.py --resume
"""

import json
import logging
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse, quote

import click
import pyautogui
from playwright.sync_api import Page, sync_playwright

from capturer import capture_region as _cap_region, refine_slide_region, get_monitor_under_mouse

QUEUE_FILE  = Path(__file__).parent / "queue_state.json"
CAPTURES_DIR = Path(__file__).parent / "captures"
LOG_FILE    = Path(__file__).parent / "crawler.log"
CDP_URL     = "http://127.0.0.1:9222"


# ── Logging setup ─────────────────────────────────────────────────────────────

def _setup_logging() -> None:
    """Write all print/click.echo output to crawler.log in addition to stdout."""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
        ],
    )


class _Tee:
    """Redirect writes to both the original stream and the log file."""
    def __init__(self, stream, log_path: Path):
        self._stream = stream
        self._log = open(log_path, "a", encoding="utf-8", buffering=1)

    def write(self, data):
        self._stream.write(data)
        self._stream.flush()
        self._log.write(data)
        self._log.flush()

    def flush(self):
        self._stream.flush()
        self._log.flush()

    def isatty(self):
        return False


# ── Queue / Resume state ──────────────────────────────────────────────────────

def load_queue() -> dict:
    if QUEUE_FILE.exists():
        with open(QUEUE_FILE) as f:
            return json.load(f)
    return {"done": [], "failed": [], "root_url": None}


def save_queue(state: dict) -> None:
    with open(QUEUE_FILE, "w") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def mark_done(state: dict, key: str) -> None:
    if key not in state["done"]:
        state["done"].append(key)
    state["failed"] = [e for e in state["failed"] if e["key"] != key]
    save_queue(state)


def mark_failed(state: dict, key: str, reason: str) -> None:
    entry = {"key": key, "reason": reason}
    state["failed"] = [e for e in state["failed"] if e["key"] != key]
    state["failed"].append(entry)
    save_queue(state)


def is_done(state: dict, key: str) -> bool:
    return key in state["done"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def slugify(name: str) -> str:
    """Convert folder/file name to safe filesystem name."""
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r'\s+', "_", name)
    return name


def wait_for_login(page: Page) -> None:
    """Detect Microsoft login page and wait for user to log in manually."""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception:
        pass

    login_indicators = [
        'input[type="email"]',
        'input[name="loginfmt"]',
        '#i0116',
        '[data-testid="login-page"]',
    ]
    is_login = any(
        page.locator(sel).count() > 0
        for sel in login_indicators
    )
    if is_login:
        click.echo("\n" + "=" * 60)
        click.echo("  PHÁT HIỆN TRANG ĐĂNG NHẬP MICROSOFT")
        click.echo("  Hãy đăng nhập vào SharePoint trong Chrome.")
        click.echo("  Sau khi đăng nhập xong, quay lại đây và nhấn Enter.")
        click.echo("=" * 60)
        input("  >>> Nhấn Enter sau khi đăng nhập xong: ")
        page.wait_for_load_state("networkidle", timeout=30000)


def wait_for_sharepoint(page: Page, timeout: int = 15000) -> None:
    """Wait until SharePoint file list is loaded."""
    try:
        page.wait_for_load_state("load", timeout=timeout)
    except Exception:
        pass
    # Wait for icons to appear (most reliable signal that list has rendered)
    try:
        page.wait_for_selector('[data-icon-name]', timeout=timeout)
    except Exception:
        pass
    wait_for_login(page)
    time.sleep(0.3)


KNOWN_FILE_EXTENSIONS = {
    "pptx", "ppt", "pdf", "docx", "doc", "xlsx", "xls",
    "jpg", "jpeg", "png", "gif", "mp4", "zip", "rar", "txt", "csv",
}

FOLDER_ICON_KEYWORDS = ["folder"]
FILE_ICON_KEYWORDS = ["powerpoint", "word", "excel", "pdf", "document", "spreadsheet", "presentation"]


def _classify_item(name: str, icon_name: str) -> str:
    """Return 'folder' or 'file' based on icon name and filename heuristics."""
    icon_lower = (icon_name or "").lower()
    if any(k in icon_lower for k in FOLDER_ICON_KEYWORDS):
        return "folder"
    if any(k in icon_lower for k in FILE_ICON_KEYWORDS):
        return "file"
    # Fallback: known file extension
    ext = name.rsplit(".", 1)[-1].lower().strip() if "." in name else ""
    if ext in KNOWN_FILE_EXTENSIONS:
        return "file"
    return "folder"


def _is_group_header(name: str) -> bool:
    skip = ["(UTC", "Pacific Time", "Eastern Time", "Central Time", "Mountain Time"]
    return any(p in name for p in skip)


def get_folder_items(page: Page) -> list[dict]:
    """
    Return list of items (folders + files) in the current SharePoint view.
    Uses JavaScript to extract items directly from DOM for robustness.
    """
    # Wait for icons (fastest signal that list has rendered)
    try:
        page.wait_for_selector('[data-icon-name]', timeout=10000)
    except Exception:
        pass

    # JavaScript-based extraction
    js_result = page.evaluate("""
    () => {
        const FOLDER_ICONS = ['folder'];
        const FILE_ICONS = ['powerpoint', 'word', 'excel', 'pdf', 'document', 'spreadsheet', 'presentation', 'onenote'];
        const KNOWN_EXTS = ['pptx','ppt','pdf','docx','doc','xlsx','xls','jpg','png','mp4','zip','txt','csv'];

        function classifyByIcon(iconName) {
            const low = (iconName || '').toLowerCase();
            if (FOLDER_ICONS.some(k => low.includes(k))) return 'folder';
            if (FILE_ICONS.some(k => low.includes(k))) return 'file';
            return null;
        }
        function classifyByName(name) {
            const ext = name.split('.').pop().toLowerCase().trim();
            return KNOWN_EXTS.includes(ext) ? 'file' : 'folder';
        }
        function isInBreadcrumb(el) {
            let cur = el;
            for (let i = 0; i < 10; i++) {
                if (!cur) break;
                const tag = cur.tagName || '';
                const role = cur.getAttribute('role') || '';
                const cls = cur.className || '';
                const aid = cur.getAttribute('data-automationid') || '';
                const aria = cur.getAttribute('aria-label') || '';
                if (
                    role === 'navigation' ||
                    aid.toLowerCase().includes('breadcrumb') ||
                    aria.toLowerCase().includes('breadcrumb') ||
                    cls.toLowerCase().includes('breadcrumb') ||
                    tag === 'NAV'
                ) return true;
                cur = cur.parentElement;
            }
            return false;
        }

        // Find the main list container (scrollable region with file rows)
        const listRoot = (
            document.querySelector('[data-automationid="list-page"]') ||
            document.querySelector('[role="grid"]') ||
            document.querySelector('[role="listbox"]') ||
            document.querySelector('.ms-DetailsList') ||
            document.querySelector('[data-automationid="DetailsListZone"]') ||
            document.body
        );

        const seen = new Set();
        const items = [];

        // Strategy A: icons with folder/file type inside list container
        const icons = listRoot.querySelectorAll('[data-icon-name]');
        for (const icon of icons) {
            if (isInBreadcrumb(icon)) continue;
            const iconName = icon.getAttribute('data-icon-name') || '';
            const itemType = classifyByIcon(iconName);
            if (!itemType) continue;

            // Walk up to row container
            let container = icon.parentElement;
            for (let i = 0; i < 10; i++) {
                if (!container || container === listRoot) break;
                const role = container.getAttribute('role') || '';
                if (role === 'row' || role === 'listitem' || container.tagName === 'TR' || container.tagName === 'LI') break;
                container = container.parentElement;
            }
            if (!container) continue;

            const nameEl = container.querySelector('span[title], a[title]');
            if (!nameEl) continue;
            const name = (nameEl.getAttribute('title') || nameEl.textContent || '').trim();
            if (!name || seen.has(name) || isInBreadcrumb(nameEl)) continue;
            seen.add(name);
            // Capture href: find the nearest <a> ancestor or descendant
            let href = '';
            let linkEl = nameEl.tagName === 'A' ? nameEl : container.querySelector('a[href]');
            if (!linkEl) {
                let cur = nameEl.parentElement;
                for (let i = 0; i < 8; i++) {
                    if (!cur) break;
                    if (cur.tagName === 'A' && cur.href) { linkEl = cur; break; }
                    cur = cur.parentElement;
                }
            }
            if (linkEl && linkEl.href && !linkEl.href.startsWith('javascript')) href = linkEl.href;
            items.push({ name, type: itemType, icon: iconName, href });
        }

        // Strategy B: span[title] inside list container
        if (items.length === 0) {
            const spans = listRoot.querySelectorAll('span[title], a[title]');
            const skipWords = new Set(['Name', 'Modified', 'Modified By', 'File Size', 'Sharing', 'Documents']);
            for (const el of spans) {
                if (isInBreadcrumb(el)) continue;
                const name = (el.getAttribute('title') || el.textContent || '').trim();
                if (!name || seen.has(name) || skipWords.has(name) || name.includes('(UTC')) continue;
                seen.add(name);
                let iconName = '';
                let href = '';
                let cur = el.parentElement;
                for (let i = 0; i < 8; i++) {
                    if (!cur) break;
                    if (!iconName) {
                        const iconEl = cur.querySelector('[data-icon-name]');
                        if (iconEl) iconName = iconEl.getAttribute('data-icon-name') || '';
                    }
                    if (!href && cur.tagName === 'A' && cur.href && !cur.href.startsWith('javascript')) {
                        href = cur.href;
                    }
                    cur = cur.parentElement;
                }
                items.push({ name, type: classifyByIcon(iconName) || classifyByName(name), icon: iconName, href });
            }
        }

        return items;
    }
    """)

    items = []
    seen = set()
    for item in (js_result or []):
        name = item.get("name", "").strip()
        if not name or _is_group_header(name) or name in seen:
            continue
        seen.add(name)
        items.append({
            "name": name,
            "type": item.get("type", "folder"),
            "href": item.get("href", ""),
        })

    return items


def sharepoint_subfolder_url(current_url: str, folder_name: str) -> str:
    """
    Build the SharePoint URL for a subfolder by appending folder_name
    to the 'id' query parameter of the current URL.
    SharePoint uses ?id=<server_relative_path> for folder navigation.
    """
    parsed = urlparse(current_url)
    params = parse_qs(parsed.query, keep_blank_values=True)
    folder_path = params.get("id", [""])[0]
    if not folder_path:
        return ""
    new_path = folder_path.rstrip("/") + "/" + folder_name
    params["id"] = [new_path]
    new_query = urlencode({k: v[0] for k, v in params.items()})
    return urlunparse(parsed._replace(query=new_query))


def navigate_to_item(page: Page, item: dict, base_url: str = "") -> None:
    """Navigate to a subfolder using SharePoint URL construction."""
    name = item["name"]
    current = base_url or page.url

    # Build URL by appending folder name to the id= parameter
    target_url = sharepoint_subfolder_url(current, name)
    if target_url:
        page.goto(target_url)
        return

    # Fallback: JS click
    prev_url = page.url
    page.evaluate(f"""
    () => {{
        const name = {repr(name)};
        const spans = document.querySelectorAll('span[title], a[title]');
        for (const span of spans) {{
            const t = (span.getAttribute('title') || span.textContent || '').trim();
            if (t !== name || span.offsetParent === null) continue;
            let el = span;
            for (let i = 0; i < 8; i++) {{
                if (!el) break;
                if (el.tagName === 'A' || el.tagName === 'BUTTON') {{ el.click(); return; }}
                el = el.parentElement;
            }}
            span.click();
        }}
    }}
    """)
    try:
        page.wait_for_function(f"window.location.href !== {repr(prev_url)}", timeout=8000)
    except Exception:
        pass




def _site_origin_base(url: str) -> tuple:
    """Return (origin, site_base_path) from any SharePoint URL.
    e.g. ('https://tenant.sharepoint.com', '/sites/SiteName')
    """
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    base = parsed.path.split("/_layouts")[0].split("/Shared%20Documents")[0].split("/Forms")[0]
    return origin, base


def get_file_guid(page: Page, server_relative_path: str) -> Optional[str]:
    """
    Get a file's GUID via SharePoint REST API, using the current page's auth session.
    Returns GUID string like 'F0023994-7374-41D0-B6AC-D2143048AAE4', or None on failure.

    Uses encodeURIComponent inside JS to handle spaces and non-ASCII chars correctly.
    Passes path as argument to avoid Python f-string / encoding issues.
    """
    try:
        result = page.evaluate(
            """
            async (path) => {
                try {
                    // Build site-level API base: origin + /sites/SITENAME
                    // window.location.pathname looks like /sites/SITENAME/Shared Documents/...
                    const parts = window.location.pathname.split('/');
                    // parts = ['', 'sites', 'SITENAME', ...]
                    const sitePath = (parts[1] === 'sites' && parts[2])
                        ? '/' + parts[1] + '/' + parts[2]
                        : '';
                    const apiBase = window.location.origin + sitePath;

                    // encodeURI encodes spaces + non-ASCII but keeps slashes intact.
                    // Single quotes must be doubled for OData string literals.
                    const safePath = path.replace(/'/g, "''");
                    const encodedPath = encodeURI(safePath);
                    const url = apiBase
                        + "/_api/web/GetFileByServerRelativeUrl('"
                        + encodedPath
                        + "')?$select=UniqueId";

                    const resp = await fetch(url, {
                        headers: { Accept: "application/json;odata=nometadata" }
                    });
                    const text = await resp.text();
                    if (!resp.ok) {
                        return { error: "HTTP " + resp.status + " | apiBase: " + apiBase + " | path: " + encodedPath + " | " + text.slice(0, 200) };
                    }
                    const data = JSON.parse(text);
                    return { guid: data.UniqueId || null };
                } catch(e) {
                    return { error: String(e) };
                }
            }
            """,
            server_relative_path,
        )
        if result and result.get("guid"):
            return result["guid"]
        if result and result.get("error"):
            print(f"    [API] {result['error'][:400]}")
        return None
    except Exception as e:
        print(f"    [API Error] {e}")
        return None


def build_office_url(guid: str, file_name: str, base_url: str, action: str = "edit") -> str:
    """Build a Doc.aspx URL from a file GUID (the correct SharePoint approach)."""
    origin, base_path = _site_origin_base(base_url)
    encoded_guid = quote("{" + guid.upper() + "}")
    encoded_name = quote(file_name)
    return (f"{origin}{base_path}/_layouts/15/Doc.aspx"
            f"?sourcedoc={encoded_guid}&file={encoded_name}&action={action}")


def build_direct_file_url(server_relative_path: str, base_url: str) -> str:
    """Build a direct file URL (origin + encoded path). Works for PDFs."""
    origin, _ = _site_origin_base(base_url)
    encoded = quote(server_relative_path, safe="/:@!$&'()*+,;=~")
    return origin + encoded


def get_server_relative_path(session_url: str, file_name: str) -> str:
    """Derive the server-relative path of a file from the session folder URL."""
    params = parse_qs(urlparse(session_url).query)
    folder_path = params.get("id", [""])[0]
    if not folder_path:
        return ""
    return folder_path.rstrip("/") + "/" + file_name


_PRESENT_LABELS = [
    'Trình bày', 'Present', 'Start Slide Show', 'From Beginning', 'Từ đầu',
    # Note: 'Slide Show' is a ribbon TAB name — do NOT include it here
]

import re as _re
_PRESENT_RE = _re.compile(
    r'Trình bày|Slideshow|Slide Show|Start Slide Show|Present|Từ đầu|From Beginning',
    _re.IGNORECASE,
)

# Use aria-label first, then innerText (avoids "Present Present" from nested splits),
# then fall back to full textContent.
_BTN_LABEL_JS = """
function btnLabel(el) {
    return (el.getAttribute('aria-label') || el.getAttribute('title') ||
            el.innerText || el.textContent || '').trim().slice(0, 80);
}
"""

# Search actual buttons BEFORE tabs — "Slide Show" is a ribbon tab, not the
# Present button. Searching buttons first ensures we click "Present ▾" (top-right).
_SCAN_BTNS_JS = ("""
() => {
    %s
    const labels = %s;
    const SELS = ['button,[role="button"],[role="menuitem"]', '[role="tab"]'];
    for (const sel of SELS) {
        for (const el of document.querySelectorAll(sel)) {
            const t = btnLabel(el);
            if (t && labels.some(l => t.includes(l))) {
                el.click();
                return t;
            }
        }
    }
    const all = [...document.querySelectorAll('button,[role="button"]')]
        .map(b => btnLabel(b)).filter(t => t.length > 1).slice(0, 20);
    return 'NOTFOUND:' + JSON.stringify(all);
}
""" % (_BTN_LABEL_JS, str(_PRESENT_LABELS).replace("'", '"')))

# Check-only — same logic, NO click
_CHECK_BTNS_JS = ("""
() => {
    %s
    const labels = %s;
    const SELS = ['button,[role="button"],[role="menuitem"]', '[role="tab"]'];
    for (const sel of SELS) {
        for (const el of document.querySelectorAll(sel)) {
            const t = btnLabel(el);
            if (t && labels.some(l => t.includes(l))) {
                return t;  // found, no click
            }
        }
    }
    const all = [...document.querySelectorAll('button,[role="button"]')]
        .map(b => btnLabel(b)).filter(t => t.length > 1).slice(0, 20);
    return 'NOTFOUND:' + JSON.stringify(all);
}
""" % (_BTN_LABEL_JS, str(_PRESENT_LABELS).replace("'", '"')))


_DISMISS_OK_JS = """
() => {
    const ok = [...document.querySelectorAll('button,input[type=button]')]
        .find(b => {
            const s = window.getComputedStyle(b);
            const vis = s.display !== 'none' && s.visibility !== 'hidden';
            return vis && /^ok$/i.test((b.textContent || b.value || '').trim());
        });
    if (ok) { ok.click(); return true; }
    return false;
}
"""


def _dismiss_ok_dialog(page) -> bool:
    """Dismiss 'no permission to edit' OK dialog — searches ALL frames (cross-origin too)."""
    for frame in page.frames:
        try:
            if frame.evaluate(_DISMISS_OK_JS):
                return True
        except Exception:
            pass
    return False


def _find_and_click_present(page, timeout: float = 30.0, do_click: bool = True):
    """
    Scan all frames (including cross-origin Office Online iframe via CDP)
    for the Present/Trình bày button.
    Returns (found: bool, label: str).
    Clicks only when do_click=True; uses check-only JS when do_click=False.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        for i, frame in enumerate(page.frames):
            try:
                # Always use check-only JS to locate the element
                result = frame.evaluate(_CHECK_BTNS_JS)
                if result and not result.startswith("NOTFOUND:"):
                    frame_hint = (frame.url or "")[:60]
                    print(f"    [Present] found in frame {i} ({frame_hint}): '{result}'")
                    if do_click:
                        # Use Playwright native click (real mouse events, not JS el.click())
                        # PowerPoint Online requires proper mouse events to start slideshow.
                        clicked = False
                        for sel in [
                            'button', '[role="button"]', '[role="menuitem"]', '[role="tab"]'
                        ]:
                            loc = frame.locator(sel).filter(
                                has_text=_re.compile(
                                    r'Trình bày|Present|Start Slide Show|From Beginning|Từ đầu',
                                    _re.IGNORECASE,
                                )
                            )
                            if loc.count() > 0:
                                loc.first.click()
                                clicked = True
                                break
                        if not clicked:
                            # Last resort: JS click
                            frame.evaluate(_SCAN_BTNS_JS)
                    return True, result
            except Exception:
                pass
        time.sleep(0.5)

    # Collect debug info from all frames
    frame_dumps: list[str] = []
    for i, frame in enumerate(page.frames):
        try:
            btns = frame.evaluate("""
            () => {
                const sels = ['button', '[role="button"]', '[role="menuitem"]', '[role="tab"]', 'a[href]'];
                const out = [];
                for (const sel of sels) {
                    for (const el of document.querySelectorAll(sel)) {
                        const t = (el.textContent||el.getAttribute('aria-label')||el.getAttribute('title')||'').trim().slice(0,50);
                        if (t.length > 1) out.push(sel + ': ' + t);
                    }
                    if (out.length > 20) break;
                }
                return out.slice(0, 20);
            }
            """)
            if btns:
                frame_url = frame.url[:60] if frame.url else "about:blank"
                frame_dumps.append(f"[frame {i} {frame_url}] {btns}")
        except Exception:
            pass
    return False, ("NOTFOUND — frames:\n  " + "\n  ".join(frame_dumps))


def _activate_chrome() -> None:
    """Bring Chrome window to foreground so pyautogui keypresses reach it."""
    subprocess.run(
        ["osascript", "-e", 'tell application "Google Chrome" to activate'],
        check=False, capture_output=True,
    )
    time.sleep(0.3)


def _get_chrome_screen_center(new_page) -> Optional[dict]:
    """Return OS screen coordinates of the Chrome window's center."""
    try:
        # Use screen.availLeft/Top to locate the monitor Chrome is on.
        # This assumes Chrome is maximized or fullscreen on that monitor.
        return new_page.evaluate("""
        () => ({
            x: Math.round(window.screen.availLeft + window.screen.availWidth / 2),
            y: Math.round(window.screen.availTop + window.screen.availHeight / 2),
        })
        """)
    except Exception:
        return None


def _get_viewport_screen_region(page) -> Optional[dict]:
    """
    Return the browser CONTENT area (viewport) in OS physical-pixel coordinates.

    Uses window.screenX/Y + outerWidth/Height vs innerWidth/Height to subtract
    the browser chrome (tab bar, address bar, toolbar).  Multiplies by
    devicePixelRatio to convert CSS pixels → physical pixels (Retina-safe).
    """
    try:
        r = page.evaluate("""
        () => {
            // Horizontal chrome (side borders — normally 0 in Chrome)
            const hChrome = Math.round((window.outerWidth - window.innerWidth) / 2);
            // Vertical chrome (tab strip + address bar + toolbar)
            const vChrome = window.outerHeight - window.innerHeight;
            return {
                left:   Math.round(window.screenX + hChrome),
                top:    Math.round(window.screenY + vChrome),
                width:  Math.round(window.innerWidth),
                height: Math.round(window.innerHeight),
            };
        }
        """)
        # Sanity: discard clearly wrong values
        if r and r["width"] > 100 and r["height"] > 100:
            return r
        return None
    except Exception:
        return None


def _get_pdf_nav_button_screen_pos(page) -> Optional[tuple]:
    """
    Find the 'Go to the next page.' button in the PDF viewer and return its
    (screen_x, screen_y) in OS physical-pixel coordinates.

    Searches all frames (handles cross-origin iframes).
    Returns None if not found.
    """
    NAV_LABELS = [
        "Go to the next page.", "Go to the next page",
        "Next page", "Next Page", "next page",
    ]
    _js = """
    (navLabels) => {
        const hChrome = Math.round((window.outerWidth - window.innerWidth) / 2);
        const vChrome = window.outerHeight - window.innerHeight;
        const vpLeft = Math.round(window.screenX + hChrome);
        const vpTop  = Math.round(window.screenY + vChrome);
        for (const el of document.querySelectorAll('button,[role="button"]')) {
            const label = (
                el.getAttribute('aria-label') || el.getAttribute('title') || ''
            ).trim();
            if (navLabels.some(l => label === l)) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) {
                    return {
                        x: vpLeft + Math.round(rect.left + rect.width  / 2),
                        y: vpTop  + Math.round(rect.top  + rect.height / 2),
                        label: label,
                    };
                }
            }
        }
        return null;
    }
    """
    try:
        for frame in page.frames:
            try:
                result = frame.evaluate(_js, NAV_LABELS)
                if result:
                    print(f"    PDF nav button '{result['label']}' @ ({result['x']},{result['y']})")
                    return result["x"], result["y"]
            except Exception:
                pass
    except Exception:
        pass
    return None


def open_pptx_and_present(
    page: Page,
    name: str,
    session_url: str = "",
) -> Optional[object]:
    """
    Open a PPTX in PowerPoint Online (edit mode), click "Trình bày", start slideshow.

    Steps:
      1. Navigate to Doc.aspx?action=edit URL.
      2. Wait up to 30s for the "Trình bày" toolbar button.
      3. Move OS mouse to Chrome window (so main.py detects the right monitor).
      4. Click "Trình bày" (or F5 fallback) → fullscreen slideshow starts.
      5. Press Home → slide 1.

    Returns the new Page object, or None on failure.
    Always uses profile: pptx-preview.
    """
    def log(step: int, msg: str) -> None:
        print(f"    [{step}] {msg}")

    try:
        base_url = session_url or page.url

        # Step 1 — Get file GUID via SharePoint REST API
        srv_path = get_server_relative_path(base_url, name)
        if not srv_path:
            print(f"  [WARN] Cannot determine server path for {name}")
            return None
        log(1, f"Getting GUID for: {name}")
        guid = get_file_guid(page, srv_path)
        if not guid:
            print(f"  [WARN] Cannot get GUID — skipping {name}")
            return None
        log(1, f"GUID: {guid}")

        # Step 2 — Open in PowerPoint Online (edit mode)
        target_url = build_office_url(guid, name, base_url, action="view")
        log(2, f"Opening: {target_url[:90]}...")
        new_page = page.context.new_page()
        try:
            new_page.goto(target_url, timeout=60000, wait_until="domcontentloaded")
        except Exception as e:
            log(2, f"Warning on goto: {e}")
            # Try to continue, sometimes the page actually loaded enough
            pass
        try:
            new_page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass
        log(2, f"Page loaded: {new_page.title()!r}")
        
        # Wait 8s for UI to settle/load scripts before interacting (User suggestion)
        log(2, "Waiting 8s for PowerPoint UI to fully load...")
        time.sleep(8.0)

        # Step 3 — Bring tab front, make Chrome fullscreen (hides tab bar + address bar
        #          so the slide fills more of the screen for better capture quality),
        #          then dismiss any dialog and wait for the Present button.
        new_page.bring_to_front()
        _activate_chrome()
        time.sleep(0.5)

        # macOS Chrome fullscreen: Ctrl+Cmd+F (equivalent to the green button / F11 on Windows).
        # This maximises screen real estate for the slide and simplifies region detection.
        pyautogui.hotkey("ctrl", "command", "f")
        log(3, "Requesting Chrome fullscreen (Ctrl+Cmd+F)…")
        time.sleep(1.5)  # wait for macOS fullscreen transition animation

        win = _get_chrome_screen_center(new_page)
        if win:
            pyautogui.moveTo(win["x"], win["y"])
            time.sleep(0.2)
            log(3, f"OS mouse → screen ({win['x']}, {win['y']})")

        log(4, "Waiting for PowerPoint Online + dismissing any dialogs (up to 30s)...")
        found, label = False, ""
        deadline = time.time() + 30.0
        while time.time() < deadline:
            # Dismiss dialog if present — searches ALL frames (dialog is in Office iframe)
            if _dismiss_ok_dialog(new_page):
                log(4, "Dismissed 'no permission to edit' dialog ✓")
                time.sleep(0.5)

            # Try to find Present button across all frames
            for i, frame in enumerate(new_page.frames):
                try:
                    result = frame.evaluate(_CHECK_BTNS_JS)
                    if result and not result.startswith("NOTFOUND:"):
                        frame_hint = (frame.url or "")[:60]
                        log(4, f"Present button ready in frame {i} ({frame_hint}): '{result}'")
                        # Strategy 1: Try Ribbon Navigation (Slide Show > From Beginning) - MOST RELIABLE
                        try:
                            ribbon_tab = frame.locator('[role="tab"]').filter(
                                has_text=_re.compile(r'^Slide Show$|^Trình chiếu$', _re.IGNORECASE)
                            )
                            if ribbon_tab.count() > 0 and ribbon_tab.first.is_visible():
                                log(4, "Found 'Slide Show' tab, attempting Ribbon navigation...")
                                try:
                                    ribbon_tab.first.click()
                                    time.sleep(5.0)
                                    
                                    start_btn = frame.locator('button, [role="button"]').filter(
                                        has_text=_re.compile(r'From Beginning|Từ đầu', _re.IGNORECASE)
                                    )
                                    if start_btn.count() > 0 and start_btn.first.is_visible():
                                        btn = start_btn.first
                                        is_disabled = btn.get_attribute("aria-disabled") == "true"
                                        if is_disabled:
                                            log(4, "'From Beginning' is disabled — right-clicking slide to activate...")
                                            # Right-click on the slide area to wake up PowerPoint Online,
                                            # then Esc to close the context menu. This causes the app to
                                            # re-evaluate button states and enable "From Beginning".
                                            _activate_chrome()
                                            time.sleep(0.3)
                                            vp = _get_viewport_screen_region(new_page)
                                            if vp:
                                                _cx = vp["left"] + vp["width"] // 2
                                                _cy = vp["top"] + vp["height"] // 2
                                            else:
                                                _cx, _cy = 900, 600
                                            pyautogui.click(_cx, _cy, button='right')
                                            time.sleep(1.0)
                                            pyautogui.press("esc")
                                            time.sleep(1.0)

                                            # Re-click the Slide Show tab (context menu may have closed it)
                                            try:
                                                ribbon_tab.first.click()
                                                time.sleep(2.0)
                                            except Exception:
                                                pass

                                            # Poll for button to become enabled (up to 10s)
                                            for _wait_i in range(10):
                                                is_disabled = btn.get_attribute("aria-disabled") == "true"
                                                if not is_disabled:
                                                    log(4, f"Button enabled after right-click + {_wait_i}s ✓")
                                                    break
                                                time.sleep(1.0)

                                            if is_disabled:
                                                log(4, "Button still disabled — falling back to editor mode capture")
                                                # Click on FIRST slide thumbnail (top of panel)
                                                try:
                                                    slide1_y = vp["top"] + 200 if vp else 200
                                                    pyautogui.click(80, slide1_y)
                                                    time.sleep(0.5)
                                                except Exception:
                                                    pass
                                                return (new_page, False)

                                        try:
                                            btn.click(force=True, timeout=5000)
                                            found, label = True, "Ribbon > From Beginning"
                                            log(5, "Clicked 'From Beginning' via Ribbon ✓")
                                            break
                                        except Exception:
                                            log(4, "Ribbon click failed after activation attempt")
                                except Exception as e:
                                    log(4, f"Ribbon click failed: {e}")
                        except Exception:
                            pass

                        # Strategy 2: Click 'Present' button (top-right corner)
                        # Native Playwright click (real mouse events)
                        for sel in ['button', '[role="button"]', '[role="menuitem"]', '[role="tab"]']:
                            loc = frame.locator(sel).filter(
                                has_text=_re.compile(
                                    r'Trình bày|Present|Start Slide Show|From Beginning|Từ đầu',
                                    _re.IGNORECASE,
                                )
                            )
                            if loc.count() > 0:
                                try:
                                    loc.first.click(timeout=1000)
                                    found, label = True, result
                                    break
                                except Exception:
                                    pass
                        
                        if not found:
                            # Fallback: JS click if native click failed or element not found by locator
                            log(4, f"Native click failed for '{result}', trying JS click...")
                            try:
                                frame.evaluate(_SCAN_BTNS_JS)
                                found, label = True, result + " (JS Click)"
                            except Exception as e:
                                log(4, f"JS click failed: {e}")

                        if found:
                            break
                except Exception:
                    pass
            if found:
                break
            time.sleep(0.5)

        if found:
            log(5, f"Clicked Present button: '{label}' ✓")
            new_page.bring_to_front()
            _activate_chrome()
        else:
            log(5, f"Present button not found — F5 fallback")
            pyautogui.press("f5")

        # Step 6 — Wait for slideshow to load (5s), AGGRESSIVELY dismissing
        #          any "no permission to edit" dialog that pops up after Present click.
        #          That dialog causes fullscreen to exit if not dismissed immediately.
        log(6, "Waiting 5s for slideshow, dismissing any late dialogs...")
        for tick in range(10):  # 10 × 0.5s = 5s
            dismissed = _dismiss_ok_dialog(new_page)
            if dismissed:
                log(6, f"  Dismissed late dialog at t={tick * 0.5:.1f}s — re-activating Chrome")
                new_page.bring_to_front()
                _activate_chrome()
            time.sleep(0.5)

        # Step 7 — Verify the slideshow actually started by checking fullscreen API.
        # For old .ppt files, "From Beginning" may appear clickable but not work.
        time.sleep(1.0)
        try:
            is_fullscreen = new_page.evaluate("document.fullscreenElement !== null")
        except Exception:
            is_fullscreen = False

        if not is_fullscreen:
            log(7, "Slideshow did NOT enter fullscreen — falling back to editor mode")
            _activate_chrome()
            vp = _get_viewport_screen_region(new_page)
            if vp:
                # Click on the FIRST slide thumbnail (top of panel, ~200px from top).
                # Clicking at the center would land on slide 3+.
                slide1_y = vp["top"] + 200
                pyautogui.click(80, slide1_y)
                time.sleep(0.5)
                log(7, f"Editor mode: clicked slide 1 thumbnail at (80, {slide1_y})")
            return (new_page, False)

        # Slideshow is running. Click to give browser keyboard focus,
        # then Left arrows to return to slide 1.
        _activate_chrome()
        vp = _get_viewport_screen_region(new_page)
        if vp:
            cx = vp["left"] + vp["width"] // 2
            cy = vp["top"] + vp["height"] // 2
            log(7, f"Clicking center ({cx},{cy}) to give slideshow keyboard focus...")
            pyautogui.click(cx, cy)
            time.sleep(1.5)
            pyautogui.press("left")
            time.sleep(0.8)
            pyautogui.press("left")
            time.sleep(0.8)
            log(7, "Focus acquired, returned to slide 1 ✓")
        else:
            log(7, "Chrome activated, slideshow ready ✓")

        return (new_page, True)

    except Exception as e:
        print(f"  [WARN] Cannot open {name}: {e}")
        return None



def run_capture(
    output_name: str,
    profile: str = "pptx-preview",
    skip_present: bool = False,
    max_slides: int = 200,
    region: Optional[dict] = None,
    nav_x: Optional[int] = None,
    nav_y: Optional[int] = None,
    delay: float = 0.5,
    nav_key: Optional[str] = None,
    same_count: int = 10,
) -> bool:
    """Run capture tool in the same terminal (blocking)."""
    main_py = str(Path(__file__).parent / "main.py")
    cmd = [
        sys.executable, main_py,
        "capture",
        "--name", output_name,
        "--profile", profile,
        "--auto-close",
        "--force",
        "--delay", str(delay),
        "--skip-present",
        "--max-slides", str(max_slides),
        "--same-count", str(same_count),
    ]
    if region:
        cmd += ["--region", f"{region['left']},{region['top']},{region['width']},{region['height']}"]
    if nav_x is not None and nav_y is not None:
        cmd += ["--nav-x", str(nav_x), "--nav-y", str(nav_y)]
    if nav_key:
        cmd += ["--nav-key", nav_key]

    try:
        result = subprocess.run(cmd, cwd=str(Path(__file__).parent))
        return result.returncode == 0
    except Exception as e:
        print(f"  [ERROR] Capture failed: {e}")
        return False


# ── Core crawler ──────────────────────────────────────────────────────────────

_CLOSE_TOC_JS = """
() => {
    // Step 1: find the "Table of contents" heading element
    let tocEl = null;
    for (const el of document.querySelectorAll('*')) {
        const own = [...el.childNodes]
            .filter(n => n.nodeType === 3)
            .map(n => n.textContent.trim()).join(' ').trim();
        const aria = (el.getAttribute('aria-label') || '').trim();
        if (/^table of contents$/i.test(own) || /^table of contents$/i.test(aria)) {
            tocEl = el;
            break;
        }
    }
    if (!tocEl) return null;

    // Step 2: walk UP from the heading to find the panel container,
    //         then find a <button> with close/× label INSIDE that panel only.
    let panel = tocEl.parentElement;
    for (let i = 0; i < 8; i++) {
        if (!panel || panel === document.body) break;
        const btns = [...panel.querySelectorAll('button,[role="button"]')];
        const closeBtn = btns.find(b => {
            const label = (
                b.getAttribute('aria-label') || b.getAttribute('title') || b.innerText || ''
            ).trim();
            return /^close$/i.test(label) || label === '×' || label === '✕';
        });
        if (closeBtn) {
            closeBtn.click();
            return 'toc-close: ' + (closeBtn.getAttribute('aria-label') || closeBtn.innerText);
        }
        panel = panel.parentElement;
    }
    return null;
}
"""


def _close_toc_panel(page) -> bool:
    """Close the Table of Contents / sidebar panel in PDF viewer (all frames)."""
    for frame in page.frames:
        try:
            result = frame.evaluate(_CLOSE_TOC_JS)
            if result:
                print(f"    [TOC] Closed panel via: '{result}'")
                return True
        except Exception:
            pass
    return False


def _open_pdf_viewer(page: Page, file_item: dict, session_url: str) -> Optional[object]:
    """
    Open a PDF file in SharePoint's PDF viewer (new tab).
    Uses the href from the file listing (contains SharePoint-auth params).
    Doc.aspx does NOT support PDFs; it redirects to the folder page.
    Returns the new Page if successful, None otherwise.
    """
    name = file_item["name"]
    # Use listing href — SharePoint generates an authenticated URL for each file
    href = file_item.get("href", "")
    if not href:
        # Fallback: direct server path with ?csf=1&web=1 (SharePoint sharing params)
        srv_path = get_server_relative_path(session_url, name)
        if srv_path:
            href = build_direct_file_url(srv_path, session_url) + "?csf=1&web=1"

    if not href:
        print(f"  [WARN] Cannot get URL for PDF {name}")
        return None

    print(f"    → Opening PDF: {href[:80]}...")
    try:
        new_page = page.context.new_page()
        try:
            new_page.goto(href, timeout=60000, wait_until="domcontentloaded")
        except Exception as e:
            print(f"    [WARN] Warning on PDF goto: {e}")
            pass
        try:
            new_page.wait_for_load_state("load", timeout=15000)
        except Exception:
            pass
        time.sleep(2)

        # Close Table of Contents panel if open (it covers slide content)
        time.sleep(1.0)
        _close_toc_panel(new_page)
        time.sleep(0.3)

        new_page.bring_to_front()
        _activate_chrome()
        time.sleep(0.5)

        # Focus content
        win = _get_chrome_screen_center(new_page)
        if win:
            pyautogui.click(win["x"], win["y"])
            time.sleep(0.2)

        pyautogui.press("home")
        time.sleep(0.5)
        return new_page
    except Exception as e:
        print(f"  [WARN] Cannot open PDF {name}: {e}")
        return None


def crawl_session_folder(
    page: Page,
    session_url: str,
    session_name: str,
    subject_name: str,
    state: dict,
    max_slides: int = 200,
    delay: float = 0.5,
) -> None:
    """
    Process one session folder (buoi hoc): find PPTX + PDF files, open and capture each.
    """
    print(f"\n  Session: {session_name}" if session_name else "\n  (files directly in subject folder)")
    page.goto(session_url)
    wait_for_sharepoint(page)

    items = get_folder_items(page)
    capture_files = [
        i for i in items
        if i["type"] == "file"
        and i["name"].lower().rsplit(".", 1)[-1] in ("pptx", "ppt", "pdf", "mp4")
    ]

    if not capture_files:
        print(f"    No capturable files (PPTX/PDF/MP4) found.")
        return

    for file_item in capture_files:
        fname = file_item["name"]
        stem = Path(fname).stem
        if session_name:
            output_name = f"{slugify(subject_name)}/{slugify(session_name)}/{slugify(stem)}"
            output_dir = CAPTURES_DIR / slugify(subject_name) / slugify(session_name) / slugify(stem)
        else:
            output_name = f"{slugify(subject_name)}/{slugify(stem)}"
            output_dir = CAPTURES_DIR / slugify(subject_name) / slugify(stem)
        queue_key = output_name
        ext = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""

        # ── MP4 video: different skip/capture logic ───────────────────────────
        if ext == "mp4":
            mp4_path = output_dir / f"{slugify(stem)}.mp4"
            if is_done(state, queue_key) and mp4_path.exists() and mp4_path.stat().st_size > 100_000:
                size_mb = mp4_path.stat().st_size / (1024 * 1024)
                print(f"    [SKIP] {fname} (video {size_mb:.1f}MB already captured)")
                continue

            print(f"    Recording [MP4]: {fname}")

            if page.url != session_url:
                try:
                    page.goto(session_url, timeout=60000)
                    wait_for_sharepoint(page)
                except Exception as e:
                    print(f"    [WARN] Failed to navigate to session folder: {e}")

            # Build the SharePoint video URL (same as PPTX: get GUID → Doc.aspx)
            try:
                srv_path = get_server_relative_path(session_url, fname)
                guid = get_file_guid(page, srv_path) if srv_path else None
                if guid:
                    video_url = build_office_url(guid, fname, session_url, action="view")
                    print(f"    [1] GUID: {guid}")
                else:
                    # Fallback: direct URL
                    origin, base_path = _site_origin_base(session_url)
                    video_url = f"{origin}{srv_path}" if srv_path else None
            except Exception as e:
                print(f"    [ERROR] Cannot build video URL: {e}")
                video_url = None

            if not video_url:
                print("    [ERROR] Cannot build video URL")
                mark_failed(state, queue_key, "Cannot build video URL")
                continue

            from video_capture import capture_video
            output_dir.mkdir(parents=True, exist_ok=True)
            success = capture_video(page, video_url, mp4_path)

            if success:
                mark_done(state, queue_key)
                print(f"    Done: {output_name}")
            else:
                mark_failed(state, queue_key, "Video capture failed")
                print(f"    FAILED: {output_name}")
            continue

        # ── PPTX / PDF: slide-based capture ───────────────────────────────────
        existing_slides = list(output_dir.glob("slide_*.png")) if output_dir.exists() else []

        if is_done(state, queue_key) and len(existing_slides) >= 3:
            print(f"    [SKIP] {fname} ({len(existing_slides)} slides already captured)")
            continue

        if len(existing_slides) >= 3:
            print(f"    [SKIP] {fname} ({len(existing_slides)} slides exist)")
            mark_done(state, queue_key)
            continue

        if existing_slides:
            print(f"    [REDO] {fname} (only {len(existing_slides)} slides, redoing)")
            state["done"] = [d for d in state["done"] if d != queue_key]
            save_queue(state)

        print(f"    Capturing [{ext.upper()}]: {fname}")

        # Navigate back to session folder if needed
        if page.url != session_url:
            try:
                page.goto(session_url, timeout=60000)
                wait_for_sharepoint(page)
            except Exception as e:
                print(f"    [WARN] Failed to navigate to session folder: {e}")
                # We can try to continue anyway, it might work if we have the file URLs already

        # ── Open file and prepare for capture ─────────────────────────────────
        capture_region: Optional[dict] = None
        pdf_nav_x: Optional[int] = None
        pdf_nav_y: Optional[int] = None
        pdf_nav_key: Optional[str] = None
        in_slideshow = True

        if ext == "pdf":
            opened_page = _open_pdf_viewer(page, file_item, session_url)
            profile = "pdf-viewer"
            if opened_page:
                # Wait for render + fullscreen transition
                time.sleep(0.5)

                # Use mss to detect monitor resolution directly (avoid JS/DPR issues)
                raw_region = get_monitor_under_mouse()

                if raw_region:
                    print(f"    [Debug] Monitor detected: {raw_region}")
                    # Save debug snapshot
                    try:
                        from capturer import capture_region as _mss_cap
                        _dbg = _mss_cap(raw_region)
                        _dbg.save(str(Path(__file__).parent / "debug_pdf_viewport.png"))
                    except Exception:
                        pass

                    # Increase skip_top to avoid Chrome UI + SharePoint header (in non-fullscreen mode)
                    # Retina Chrome UI ~160-180px. Set 180 to be safe.
                    capture_region = refine_slide_region(raw_region, skip_top=150, skip_bottom=50)
                    print(f"    Auto-detected PDF region: {capture_region['width']}x{capture_region['height']}"
                          f" @ ({capture_region['left']},{capture_region['top']})"
                          + (" [refined]" if capture_region != raw_region else " [viewport]"))
                else:
                    print("    ⚠ Could not auto-detect PDF region (mouse not on monitor?)")

                nav_pos = _get_pdf_nav_button_screen_pos(opened_page)
                if nav_pos and isinstance(nav_pos, dict):
                    pdf_nav_x, pdf_nav_y = nav_pos["x"], nav_pos["y"]
                else:
                    print("    ⚠ PDF nav button not found, capture may not advance pages. Trying PageDown key.")
                    pdf_nav_key = "pagedown"
        else:
            pptx_result = open_pptx_and_present(page, fname, session_url=session_url)
            in_slideshow = True
            if pptx_result is None:
                opened_page = None
            elif isinstance(pptx_result, tuple):
                opened_page, in_slideshow = pptx_result
            else:
                opened_page = pptx_result

            profile = "pptx-preview"
            if opened_page and not in_slideshow:
                print("    [Editor Mode] Capturing slides from editor view (Down arrow navigation)")

            if opened_page:
                # Wait for the first slide to fully render before taking the detection screenshot.
                # Chrome is now fullscreen (set in open_pptx_and_present), so the viewport
                # covers the entire screen — making letterbox detection reliable.
                print("    Waiting 3s for slide to render in fullscreen...")
                time.sleep(3.0)

                # Use mss to detect the monitor under the mouse (Chrome window)
                # This handles high-DPI scaling correctly without JS math.
                raw_region = get_monitor_under_mouse()
                if raw_region:
                    print(f"    [Debug] Monitor detected: {raw_region}")
                    # Save debug snapshot so we can verify what refine_slide_region sees
                    try:
                        from capturer import capture_region as _mss_cap
                        _dbg = _mss_cap(raw_region)
                        _dbg.save(str(Path(__file__).parent / "debug_pptx_viewport.png"))
                    except Exception:
                        pass

                    capture_region = refine_slide_region(raw_region, skip_top=0, skip_bottom=0)
                    print(f"    Auto-detected PPTX region: {capture_region['width']}x{capture_region['height']}"
                          f" @ ({capture_region['left']},{capture_region['top']})"
                          + (" [refined]" if capture_region != raw_region else " [viewport]"))
                else:
                    print("    ⚠ Could not auto-detect PPTX region, using full monitor")

        if not opened_page:
            mark_failed(state, queue_key, "Could not open in browser")
            continue

        print(f"    Profile: {profile}")

        # Ensure Chrome is the foreground app right before capture starts
        _activate_chrome()

        # Run capture
        _local_ext = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
        if _local_ext == "pdf":
            nav_key_to_use = "pagedown"
        elif not in_slideshow:
            nav_key_to_use = "down"
        else:
            nav_key_to_use = "right"
        
        # Editor mode: stop quickly when screen doesn't change (3 tries vs 10)
        _same_count = 3 if (not in_slideshow) else 10
        success = run_capture(output_name, profile=profile, skip_present=True,
                              max_slides=max_slides, region=capture_region,
                              nav_x=pdf_nav_x, nav_y=pdf_nav_y, nav_key=nav_key_to_use,
                              delay=delay, same_count=_same_count)

        # Exit Chrome fullscreen (PPTX and PDF both use it; restore normal window)
        try:
            pyautogui.hotkey("ctrl", "command", "f")
            time.sleep(0.8)
        except Exception:
            pass

        # Close the file tab
        try:
            opened_page.close()
        except Exception:
            pass
        time.sleep(1)
        if success:
            mark_done(state, queue_key)
            print(f"    Done: {output_name}")
        else:
            mark_failed(state, queue_key, "Capture failed")
            print(f"    FAILED: {output_name}")

        # Note: We do NOT need to goto session_url here unconditionally because
        # we open files in a new tab. The original tab stays at the session folder.
        # There is a safety check at the start of the loop just in case.


def crawl_subject_folder(
    page: Page,
    subject_url: str,
    subject_name: str,
    state: dict,
    max_slides: int = 200,
    max_sessions: int = 0,
    delay: float = 0.5,
) -> None:
    """Process one subject folder (1.NLS, 2.KNCB...): list session folders."""
    print(f"\nSubject: {subject_name}")
    page.goto(subject_url)
    wait_for_sharepoint(page)

    items = get_folder_items(page)
    session_folders = [i for i in items if i["type"] == "folder"]

    direct_files = [
        i for i in items
        if i["type"] == "file"
        and i["name"].lower().rsplit(".", 1)[-1] in ("pptx", "ppt", "pdf", "mp4")
    ]
    if direct_files:
        print(f"  Found {len(direct_files)} file(s) directly in subject folder")
        crawl_session_folder(page, subject_url, "", subject_name, state,
                             max_slides=max_slides, delay=delay)

    if not session_folders:
        if not direct_files:
            print("  No session folders found.")
        return

    if max_sessions > 0:
        session_folders = session_folders[:max_sessions]
        print(f"  [TEST] Limiting to {max_sessions} session(s)")

    for session in session_folders:
        session_name = session["name"]
        session_url = sharepoint_subfolder_url(subject_url, session_name)
        crawl_session_folder(page, session_url, session_name, subject_name, state,
                             max_slides=max_slides, delay=delay)


def crawl_root(
    page: Page,
    root_url: str,
    state: dict,
    max_slides: int = 200,
    max_sessions: int = 0,
    max_subjects: int = 0,
    delay: float = 0.5,
) -> None:
    """Start from root 'Tai lieu bai giang', list all subject folders."""
    print(f"Root: {root_url}")
    page.goto(root_url)
    print("Waiting for page load (1s)...")
    time.sleep(1.0) # Wait explicitly for list to render
    wait_for_sharepoint(page)

    items = get_folder_items(page)
    subject_folders = [i for i in items if i["type"] == "folder"]
    print(f"Found {len(subject_folders)} subject folder(s): {[s['name'] for s in subject_folders]}")

    if max_subjects > 0:
        subject_folders = subject_folders[:max_subjects]
        print(f"[TEST] Limiting to {max_subjects} subject(s)")

    for subject in subject_folders:
        subject_name = subject["name"]
        subject_url = sharepoint_subfolder_url(root_url, subject_name)
        crawl_subject_folder(page, subject_url, subject_name, state,
                             max_slides=max_slides, max_sessions=max_sessions,
                             delay=delay)

    print("\nCrawl complete!")


# ── CLI ───────────────────────────────────────────────────────────────────────

@click.command()
@click.option("--url", default=None, help="URL trang 'Tai lieu bai giang' trên SharePoint.")
@click.option("--resume", is_flag=True, default=False, help="Tiếp tục từ lần chạy trước.")
@click.option("--cdp", default=CDP_URL, show_default=True, help="Chrome DevTools Protocol URL.")
@click.option("--subject", default=None, help="Chỉ xử lý folder cụ thể (vd: '1.NLS').")
@click.option("--dry-run", is_flag=True, default=False, help="Chỉ liệt kê file, không capture.")
@click.option("--test-flow", is_flag=True, default=False,
              help="Mở thử 1 PPTX + 1 PDF, kiểm tra UI đúng luồng, không capture.")
@click.option("--max-slides", default=200, show_default=True,
              help="Số slide tối đa mỗi file (dùng để test, vd: --max-slides 3).")
@click.option("--max-sessions", default=0, show_default=True,
              help="Số session folder tối đa mỗi subject (0=tất cả). Dùng để test.")
@click.option("--max-subjects", default=0, show_default=True,
              help="Số subject folder tối đa (0=tất cả). Dùng để test.")
@click.option("--delay", default=0.5, type=float, show_default=True,
              help="Giây chờ slide load (default 0.5s).")
def main(url: Optional[str], resume: bool, cdp: str, subject: Optional[str],
         dry_run: bool, test_flow: bool,
         max_slides: int, max_sessions: int, max_subjects: int, delay: float):
    """
    SharePoint Auto Crawler - tự động chụp toàn bộ PPTX từ SharePoint.

    Trước khi chạy, mở Chrome với remote debugging:
      open -a "Google Chrome" --args --remote-debugging-port=9222
    """
    # ── Tee stdout+stderr to crawler.log ──────────────────────────────────────
    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tee_stdout = _Tee(sys.stdout, LOG_FILE)
    tee_stderr = _Tee(sys.stderr, LOG_FILE)
    sys.stdout = tee_stdout  # type: ignore[assignment]
    sys.stderr = tee_stderr  # type: ignore[assignment]
    print(f"\n{'='*60}")
    print(f"  Crawler started at {run_ts}")
    print(f"  Log: {LOG_FILE}")
    print(f"{'='*60}\n")

    state = load_queue()

    if resume and state.get("root_url"):
        url = state["root_url"]
        click.echo(f"Resuming from: {url}")
        click.echo(f"Already done: {len(state['done'])} files")
        click.echo(f"Failed: {len(state['failed'])} files")
    elif not url:
        click.echo("Error: --url is required for first run.", err=True)
        raise SystemExit(1)
    else:
        state["root_url"] = url
        save_queue(state)

    click.echo(f"\nConnecting to Chrome at {cdp}...")
    click.echo("Nếu chưa mở Chrome, hãy chạy lệnh sau:")
    click.echo("  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome \\")
    click.echo("    --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-crawl\n")

    with sync_playwright() as pw:
        try:
            browser = pw.chromium.connect_over_cdp(cdp)
        except Exception as e:
            click.echo(f"\nKhông kết nối được Chrome: {e}", err=True)
            click.echo("\nMở Chrome với remote debugging (PHẢI có --user-data-dir):")
            click.echo("  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome \\")
            click.echo("    --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-crawl")
            click.echo("\nSau đó đăng nhập SharePoint trong Chrome, rồi chạy lại crawler.")
            raise SystemExit(1)

        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.pages[0] if context.pages else context.new_page()

        if dry_run:
            click.echo("[DRY RUN] Listing files...\n")
            try:
                _dry_run_list(page, url, subject)
            except Exception as e:
                import traceback
                click.echo(f"\n[ERROR] {e}", err=True)
                traceback.print_exc()
        elif test_flow:
            click.echo("[TEST FLOW] Opening sample files to verify UI...\n")
            try:
                _test_flow(page, url, subject)
            except Exception as e:
                import traceback
                click.echo(f"\n[ERROR] {e}", err=True)
                traceback.print_exc()
        else:
            if subject:
                # Single subject mode — build URL directly, no click needed
                subject_url = sharepoint_subfolder_url(url, subject)
                if not subject_url:
                    click.echo(f"Cannot build URL for subject '{subject}'. Check --url has ?id= param.", err=True)
                    raise SystemExit(1)
                crawl_subject_folder(page, subject_url, subject, state,
                                     max_slides=max_slides, max_sessions=max_sessions,
                                     delay=delay)
            else:
                crawl_root(page, url, state,
                           max_slides=max_slides,
                           max_sessions=max_sessions,
                           max_subjects=max_subjects,
                           delay=delay)

        click.echo("\nDone! Check captures/ folder for results.")
        failed = state.get("failed", [])
        if failed:
            click.echo(f"\nFailed files ({len(failed)}):")
            for f in failed:
                click.echo(f"  - {f['key']}: {f['reason']}")



def _test_flow(page: Page, url: str, subject_filter: Optional[str]) -> None:
    """
    Find the first PPTX and first PDF in the folder structure.
    Open each one, verify the expected UI state, then close — NO capturing.

    PPTX check : "Trình bày" toolbar button appears within 30s
    PDF  check : page loads and has content within 15s
    """
    # ── Walk folders to collect 1 PPTX + 1 PDF sample ────────────────────────
    page.goto(url, timeout=30000)
    wait_for_sharepoint(page)

    subjects = [i for i in get_folder_items(page) if i["type"] == "folder"]
    sample: dict = {"pptx": None, "pdf": None, "pptx_url": "", "pdf_url": ""}

    for subj in subjects:
        if subject_filter and subj["name"] != subject_filter:
            continue
        subj_url = sharepoint_subfolder_url(url, subj["name"])
        page.goto(subj_url)
        wait_for_sharepoint(page)

        for sess in [i for i in get_folder_items(page) if i["type"] == "folder"]:
            sess_url = sharepoint_subfolder_url(subj_url, sess["name"])
            page.goto(sess_url)
            wait_for_sharepoint(page)

            for f in get_folder_items(page):
                if f["type"] != "file":
                    continue
                ext = f["name"].rsplit(".", 1)[-1].lower() if "." in f["name"] else ""
                if ext in ("pptx", "ppt") and not sample["pptx"]:
                    sample["pptx"] = f
                    sample["pptx_url"] = sess_url
                if ext == "pdf" and not sample["pdf"]:
                    sample["pdf"] = f
                    sample["pdf_url"] = sess_url
            if sample["pptx"] and sample["pdf"]:
                break

            page.goto(subj_url)
            wait_for_sharepoint(page)
        if sample["pptx"] and sample["pdf"]:
            break

    if not sample["pptx"] and not sample["pdf"]:
        click.echo("No PPTX or PDF files found to test.")
        return

    results = []

    # ── Test PPTX ─────────────────────────────────────────────────────────────
    if sample["pptx"]:
        fname = sample["pptx"]["name"]
        sess_url = sample["pptx_url"]
        click.echo(f"[PPTX] Testing: {fname}")

        srv_path = get_server_relative_path(sess_url, fname)
        click.echo(f"       Server path: {srv_path}")
        guid = get_file_guid(page, srv_path) if srv_path else None
        if not guid:
            click.echo("       ❌ Cannot get GUID from SharePoint API")
            results.append(("PPTX", fname, "FAIL — no GUID"))
        else:
            target_url = build_office_url(guid, fname, sess_url, action="view")
            click.echo(f"       GUID: {guid}")
            click.echo(f"       URL: {target_url[:90]}...")

        if guid:
            try:
                tab = page.context.new_page()
                tab.goto(target_url, timeout=30000)
                try:
                    tab.wait_for_load_state("domcontentloaded", timeout=15000)
                except Exception:
                    pass
                click.echo(f"       Page: {tab.title()!r}")

                # Poll: dismiss dialog (in Office iframe!) + find Present button (up to 30s)
                tab.bring_to_front()
                _activate_chrome()
                found, label = False, ""
                deadline2 = time.time() + 30.0
                while time.time() < deadline2:
                    if _dismiss_ok_dialog(tab):
                        click.echo("       → Dismissed 'no permission' dialog (Office iframe)")
                        time.sleep(0.5)
                    for i, frame in enumerate(tab.frames):
                        try:
                            result = frame.evaluate(_CHECK_BTNS_JS)
                            if result and not result.startswith("NOTFOUND:"):
                                found, label = True, result
                                break
                        except Exception:
                            pass
                    if found:
                        break
                    time.sleep(0.5)

                # Now click Present (if found)
                if found:
                    # Actually click it with Playwright native click
                    for frame in tab.frames:
                        try:
                            for sel in ['button', '[role="button"]', '[role="menuitem"]']:
                                loc = frame.locator(sel).filter(
                                    has_text=_re.compile(
                                        r'Trình bày|Present|Start Slide Show|Từ đầu',
                                        _re.IGNORECASE,
                                    )
                                )
                                if loc.count() > 0:
                                    loc.first.click()
                                    break
                        except Exception:
                            pass

                if not found:
                    click.echo(f'       ❌ Present button NOT found. Title: {tab.title()!r}')
                    click.echo(f"       Debug: {label}")
                    results.append(("PPTX", fname, "FAIL — button not found"))
                else:
                    click.echo(f'       ✅ Clicked Present button: "{label}"')
                    # Bring the tab to OS foreground so user can see the slideshow
                    tab.bring_to_front()
                    _activate_chrome()
                    click.echo("       Waiting 3s for slideshow to load...")
                    time.sleep(3)
                    # Screenshot: what does the tab look like right now?
                    shot1 = str(Path(__file__).parent / "test_after_present.png")
                    try:
                        tab.screenshot(path=shot1, full_page=False)
                        click.echo(f"       📸 Screenshot saved: {shot1}")
                    except Exception as se:
                        click.echo(f"       ⚠ Screenshot failed: {se}")

                    try:
                        # Use pyautogui for OS-level keypresses (works in slideshow mode)
                        pyautogui.press("home")
                        time.sleep(1.0)

                        # Auto-detect PPTX crop region (no toolbar at top/bottom → skip=0)
                        raw_region_pptx = _get_viewport_screen_region(tab)
                        if raw_region_pptx:
                            refined_pptx = refine_slide_region(raw_region_pptx, skip_top=0, skip_bottom=0)
                            r_icon = "refined" if refined_pptx != raw_region_pptx else "viewport"
                            click.echo(f"       📐 Viewport: {raw_region_pptx['width']}x{raw_region_pptx['height']}"
                                       f" @ ({raw_region_pptx['left']},{raw_region_pptx['top']})")
                            click.echo(f"       📐 Slide region ({r_icon}): "
                                       f"{refined_pptx['width']}x{refined_pptx['height']}"
                                       f" @ ({refined_pptx['left']},{refined_pptx['top']})")

                        # Screenshot after Home
                        shot2 = str(Path(__file__).parent / "test_slide1.png")
                        try:
                            tab.screenshot(path=shot2, full_page=False)
                            click.echo(f"       📸 Slide 1 screenshot: {shot2}")
                        except Exception:
                            pass
                        click.echo("       Navigating slides 1 → 2 → 3 (1.5s each)...")
                        for slide_n in range(2, 5):
                            pyautogui.press("right")
                            time.sleep(1.5)
                            shot = str(Path(__file__).parent / f"test_slide{slide_n}.png")
                            try:
                                tab.screenshot(path=shot, full_page=False)
                            except Exception:
                                pass
                            click.echo(f"       → slide {slide_n}  📸 {shot}")
                        click.echo("       ✅ Slide navigation working")
                        results.append(("PPTX", fname, "PASS"))
                    except Exception as e:
                        click.echo(f"       ⚠ Navigation error: {e}")
                        results.append(("PPTX", fname, "PASS (clicked, nav unverified)"))
                    # Exit slideshow
                    try:
                        pyautogui.press("escape")
                        time.sleep(1.0)
                    except Exception:
                        pass

                tab.close()
            except Exception as e:
                click.echo(f"       ❌ Error: {e}")
                results.append(("PPTX", fname, f"ERROR: {e}"))

    # ── Test PDF ──────────────────────────────────────────────────────────────
    if sample["pdf"]:
        fname = sample["pdf"]["name"]
        sess_url = sample["pdf_url"]
        click.echo(f"\n[PDF ] Testing: {fname}")

        # Build direct URL (csf=1&web=1 opens SharePoint PDF viewer with auth)
        srv_path = get_server_relative_path(sess_url, fname)
        href = (build_direct_file_url(srv_path, sess_url) + "?csf=1&web=1") if srv_path else ""
        click.echo(f"       URL: {href[:80]}...")

        if not href:
            click.echo("       ❌ Cannot build URL")
            results.append(("PDF", fname, "FAIL — no URL"))
        else:
            try:
                tab = page.context.new_page()
                tab.goto(href, timeout=30000)
                try:
                    tab.wait_for_load_state("load", timeout=20000)
                except Exception:
                    pass
                time.sleep(2)

                title = tab.title()
                bad = ("All Documents" in title or "AllItems" in title
                       or "Access required" in title or "Sign in" in title)
                if bad:
                    click.echo(f"       ❌ Bad page — title: '{title}'")
                    results.append(("PDF", fname, "FAIL"))
                    tab.close()
                else:
                    click.echo(f"       ✅ PDF viewer loaded — title: '{title}'")
                    click.echo(f"       Current URL: {tab.url[:100]}")

                    # Close TOC panel then bring to front
                    time.sleep(1.0)
                    url_before_toc = tab.url
                    closed = _close_toc_panel(tab)
                    time.sleep(0.3)
                    url_after_toc = tab.url
                    if url_after_toc != url_before_toc:
                        click.echo(f"       ⚠ TOC close caused navigation! URL changed:")
                        click.echo(f"         before: {url_before_toc[:80]}")
                        click.echo(f"         after:  {url_after_toc[:80]}")
                    else:
                        click.echo(f"       {'✅ Closed TOC panel' if closed else '⚠ TOC panel not found (may already be closed)'}")

                    tab.bring_to_front()
                    _activate_chrome()
                    time.sleep(0.5)

                    # Auto-detect capture region: viewport coords → fine-trim toolbar
                    raw_region = _get_viewport_screen_region(tab)
                    detected_region = refine_slide_region(raw_region) if raw_region else None
                    if detected_region:
                        refined = detected_region != raw_region
                        click.echo(f"       📐 Viewport: {raw_region['width']}x{raw_region['height']}"
                                   f" @ ({raw_region['left']},{raw_region['top']})")
                        click.echo(f"       📐 Slide region ({'refined' if refined else 'viewport'}): "
                                   f"{detected_region['width']}x{detected_region['height']}"
                                   f" @ ({detected_region['left']},{detected_region['top']})")
                    else:
                        click.echo("       ⚠ Region detection failed")

                    # Screenshot: Playwright viewport (full content area, for reference)
                    shot_pdf1_raw = str(Path(__file__).parent / "test_pdf_page1_raw.png")
                    try:
                        tab.screenshot(path=shot_pdf1_raw, full_page=False)
                        click.echo(f"       📸 PDF raw (Playwright viewport): {shot_pdf1_raw}")
                    except Exception:
                        pass

                    # Screenshot: mss crop with refined slide region (what capture will see)
                    shot_pdf1 = str(Path(__file__).parent / "test_pdf_page1.png")
                    if detected_region:
                        try:
                            slide_img = _cap_region(detected_region)
                            slide_img.save(shot_pdf1)
                            click.echo(f"       📸 PDF page 1 (mss {detected_region['width']}x{detected_region['height']}): {shot_pdf1}")
                        except Exception as e:
                            click.echo(f"       ⚠ mss crop failed: {e}")
                    else:
                        try:
                            tab.screenshot(path=shot_pdf1, full_page=False)
                            click.echo(f"       📸 PDF page 1: {shot_pdf1}")
                        except Exception:
                            pass

                    # Test navigation: click next-page button via Playwright DOM
                    # Use exact labels only — avoid 'next' matching 'Next steps' etc.
                    nav_js = """
                    () => {
                        // Exact pagination labels only (case-insensitive exact match or specific patterns)
                        function isNextPage(label) {
                            const l = label.trim().toLowerCase();
                            return l === 'next page' || l === 'next slide' ||
                                   l === 'suivant' || l === '下一页' ||
                                   l === 'next' || l === 'forward' ||
                                   l === 'go to the next page.' || l === 'go to the next page' ||
                                   /^next\\s*page$/i.test(l) || /^page suivante$/i.test(l) ||
                                   /^go to (the )?next page/i.test(l);
                        }
                        // Prioritize buttons with specific aria-label
                        for (const el of document.querySelectorAll('button,[role="button"]')) {
                            const ariaLabel = (el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();
                            if (ariaLabel && isNextPage(ariaLabel)) {
                                el.click();
                                return 'aria: ' + ariaLabel;
                            }
                        }
                        // Try innerText but only for short, specific labels (not panels/menus)
                        for (const el of document.querySelectorAll('button')) {
                            const txt = (el.innerText || '').trim();
                            if (txt && txt.length < 20 && isNextPage(txt)) {
                                el.click();
                                return 'txt: ' + txt;
                            }
                        }
                        // Dump all button labels for debug
                        const all = [...document.querySelectorAll('button,[role="button"]')]
                            .map(b => (b.getAttribute('aria-label')||b.innerText||'').trim().slice(0,40))
                            .filter(t => t.length > 0).slice(0, 20);
                        return 'NOTFOUND:' + JSON.stringify(all);
                    }
                    """
                    nav_found = False
                    for frame in tab.frames:
                        try:
                            res = frame.evaluate(nav_js)
                            if res and not res.startswith("NOTFOUND:"):
                                click.echo(f"       ✅ Clicked next-page: '{res}'")
                                nav_found = True
                                break
                            elif res and res.startswith("NOTFOUND:"):
                                click.echo(f"       [frame {tab.frames.index(frame)}] buttons: {res[9:]}")
                        except Exception:
                            pass

                    time.sleep(1.5)
                    shot_pdf2 = str(Path(__file__).parent / "test_pdf_page2.png")
                    if detected_region:
                        try:
                            slide_img2 = _cap_region(detected_region)
                            slide_img2.save(shot_pdf2)
                            click.echo(f"       📸 PDF page 2 (mss crop): {shot_pdf2}")
                        except Exception as e:
                            click.echo(f"       ⚠ mss crop page 2 failed: {e}")
                    else:
                        try:
                            tab.screenshot(path=shot_pdf2, full_page=False)
                            click.echo(f"       📸 PDF page 2: {shot_pdf2}")
                        except Exception:
                            pass

                    if nav_found:
                        results.append(("PDF", fname, "PASS"))
                    else:
                        click.echo("       ⚠ Next-page button not found in any frame")
                        results.append(("PDF", fname, "PASS (viewer OK, nav not found)"))

                    tab.close()
            except Exception as e:
                click.echo(f"       ❌ Error: {e}")
                results.append(("PDF", fname, f"ERROR: {e}"))

    # ── Summary ───────────────────────────────────────────────────────────────
    click.echo("\n" + "─" * 55)
    click.echo("  TEST RESULTS")
    click.echo("─" * 55)
    for ftype, fname, status in results:
        icon = "✅" if status == "PASS" else "❌"
        click.echo(f"  {icon} [{ftype}] {fname[:50]}  →  {status}")
    all_pass = all(r[2] == "PASS" for r in results)
    click.echo("─" * 55)
    if all_pass:
        click.echo("  ✅ All checks passed — ready to run full crawler!")
    else:
        click.echo("  ❌ Some checks failed — review output above before running full crawler.")
    click.echo("─" * 55)


_FILE_TYPE_LABELS = {
    "pptx": "PPTX",
    "ppt":  "PPT ",
    "pdf":  "PDF ",
    "mp4":  "MP4 ",
    "docx": "DOCX",
    "doc":  "DOC ",
    "xlsx": "XLSX",
    "xls":  "XLS ",
}

_CAPTURE_EXTS = {"pptx", "ppt", "pdf", "mp4"}


def _file_label(name: str) -> str:
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    return _FILE_TYPE_LABELS.get(ext, ext.upper()[:4].ljust(4))


def _dry_run_list(page: Page, url: str, subject_filter: Optional[str]) -> None:
    """
    Tree-view listing of all files in the SharePoint folder structure.

      📁 Subject/
        📁 Session/
          → [PPTX] filename.pptx   ← will capture (pptx-preview)
          → [PDF ] filename.pdf    ← will capture (pdf-viewer)
          · [DOCX] filename.docx   ← skipped
    """
    page.goto(url, timeout=30000)
    wait_for_sharepoint(page)

    all_items = get_folder_items(page)
    subjects = [i for i in all_items if i["type"] == "folder"]

    if not subjects:
        click.echo("No subject folders found at root URL.")
        click.echo(f"Items found ({len(all_items)}): {[i['name'] for i in all_items]}")
        return

    total: dict = {"pptx": 0, "pdf": 0, "other": 0, "sessions": 0}
    click.echo(f"Found {len(subjects)} subject folder(s):\n")

    for subj in subjects:
        if subject_filter and subj["name"] != subject_filter:
            continue

        click.echo(f"📁 {subj['name']}/")
        subj_url = sharepoint_subfolder_url(url, subj["name"])
        page.goto(subj_url)
        wait_for_sharepoint(page)

        sessions = [i for i in get_folder_items(page) if i["type"] == "folder"]
        if not sessions:
            click.echo("    (no session folders)")
            continue

        for sess in sessions:
            total["sessions"] += 1
            click.echo(f"  📁 {sess['name']}/")
            sess_url = sharepoint_subfolder_url(subj_url, sess["name"])
            page.goto(sess_url)
            wait_for_sharepoint(page)

            files = [i for i in get_folder_items(page) if i["type"] == "file"]
            if not files:
                click.echo("      (empty)")
            for f in files:
                name = f["name"]
                ext  = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                mark = "→" if ext in _CAPTURE_EXTS else "·"
                label = _file_label(name)
                click.echo(f"    {mark} [{label}] {name}")
                if ext in ("pptx", "ppt"):
                    total["pptx"] += 1
                elif ext == "pdf":
                    total["pdf"] += 1
                elif ext == "mp4":
                    total["mp4"] = total.get("mp4", 0) + 1
                else:
                    total["other"] += 1

            page.goto(subj_url)
            wait_for_sharepoint(page)

    click.echo("\n" + "─" * 55)
    click.echo(f"  Sessions : {total['sessions']}")
    click.echo(f"  PPTX/PPT : {total['pptx']}  (→ pptx-preview profile)")
    click.echo(f"  PDF      : {total['pdf']}  (→ pdf-viewer profile)")
    click.echo(f"  Other    : {total['other']}  (· skipped)")
    click.echo("─" * 55)


if __name__ == "__main__":
    main()
