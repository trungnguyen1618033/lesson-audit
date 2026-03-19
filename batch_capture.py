import json
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlparse, unquote

import click
from playwright.sync_api import sync_playwright

import video_capture
from video_capture import _is_valid_video
from crawler import (
    get_folder_items,
    get_server_relative_path,
    _site_origin_base,
    wait_for_sharepoint,
    slugify,
)

REPORT_DIR = Path("captures/_reports")
DEFAULT_OUTPUT_DIR = Path("captures/recordings")
LOG_FILE = Path("captures/_reports/batch_capture.log")


class _Tee:
    def __init__(self, stream, log_path: Path):
        log_path.parent.mkdir(parents=True, exist_ok=True)
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


def _discover_all_mp4(page, folder_url: str) -> list[dict]:
    page.goto(folder_url, timeout=60000, wait_until="domcontentloaded")
    wait_for_sharepoint(page)
    time.sleep(2)

    parsed = urlparse(folder_url)
    params = dict(p.split("=", 1) for p in parsed.query.split("&") if "=" in p)
    folder_path = unquote(params.get("id", ""))
    if not folder_path:
        click.echo("  [WARN] Cannot extract folder path from URL, falling back to DOM scroll")
        return _discover_via_scroll(page)

    click.echo(f"  Querying SharePoint REST API for folder: ...{folder_path[-40:]}")
    api_result = page.evaluate("""async (folderPath) => {
        try {
            const parts = window.location.pathname.split('/');
            const sitePath = (parts[1] === 'sites' && parts[2])
                ? '/' + parts[1] + '/' + parts[2] : '';
            const apiBase = window.location.origin + sitePath;
            const safePath = folderPath.replace(/'/g, "''");
            const url = apiBase + "/_api/web/GetFolderByServerRelativeUrl('"
                + encodeURI(safePath)
                + "')/Files?$select=Name,Length,TimeLastModified&$top=5000";
            const resp = await fetch(url, {
                headers: { Accept: 'application/json;odata=nometadata' }
            });
            if (!resp.ok) return { error: 'HTTP ' + resp.status };
            const data = await resp.json();
            return { files: data.value || [] };
        } catch(e) {
            return { error: String(e) };
        }
    }""", folder_path)

    if "error" in api_result:
        click.echo(f"  [WARN] REST API failed: {api_result['error']}, falling back to DOM scroll")
        return _discover_via_scroll(page)

    all_files = api_result.get("files", [])
    mp4_list = [
        {"name": f["Name"], "type": "file", "href": "", "size": f.get("Length", 0)}
        for f in all_files
        if f["Name"].lower().endswith(".mp4")
    ]
    click.echo(f"  REST API: {len(mp4_list)} MP4 (of {len(all_files)} total files)")
    return sorted(mp4_list, key=lambda x: x["name"])


def _discover_via_scroll(page) -> list[dict]:
    all_mp4: dict[str, dict] = {}
    prev_count = -1
    no_new = 0

    for _ in range(120):
        items = get_folder_items(page)
        for it in items:
            if it["type"] == "file" and it["name"].lower().endswith(".mp4"):
                all_mp4[it["name"]] = it

        if len(all_mp4) == prev_count:
            no_new += 1
            if no_new >= 6:
                break
        else:
            no_new = 0
        prev_count = len(all_mp4)

        page.evaluate("""() => {
            const el = document.querySelector('[data-automationid="spgrid"]');
            if (el) el.scrollTop += 800;
        }""")
        time.sleep(1.0)

    return sorted(all_mp4.values(), key=lambda x: x["name"])


def _has_video_element(video_page) -> bool:
    for frame in video_page.frames:
        try:
            if frame.evaluate("() => !!document.querySelector('video')"):
                return True
        except Exception:
            pass
    return False


def _open_video_by_click(page, folder_url: str, fname: str):
    page.goto(folder_url, timeout=60000, wait_until="domcontentloaded")
    wait_for_sharepoint(page)
    time.sleep(2)

    unique_sub = fname.rsplit("-", 1)[0] if "-" in fname else fname.rsplit(".", 1)[0]
    rect = page.evaluate("""(sub) => {
        const buttons = document.querySelectorAll('span[role="button"]');
        for (const btn of buttons) {
            const text = (btn.textContent || '').trim();
            if (text.includes(sub) && text.endsWith('.mp4')) {
                // Exclude variants like "name 1.mp4", "name 2.mp4" if exact match needed
                const r = btn.getBoundingClientRect();
                if (r.width > 0 && r.height > 0)
                    return { x: r.x + r.width / 2, y: r.y + r.height / 2, text: text };
            }
        }
        return null;
    }""", unique_sub)

    if not rect:
        for _ in range(20):
            page.mouse.wheel(0, 1500)
            time.sleep(0.8)
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

    if not rect:
        click.echo(f"    [WARN] Cannot find '{fname}' in folder DOM")
        return None

    new_pages = []
    page.context.on("page", lambda p: new_pages.append(p))

    page.mouse.click(rect["x"], rect["y"])

    deadline = time.time() + 15
    while not new_pages and time.time() < deadline:
        time.sleep(0.5)

    try:
        page.context.remove_listener("page", lambda p: None)
    except Exception:
        pass

    if not new_pages:
        click.echo("    [WARN] Click did not open a new tab")
        return None

    vp = new_pages[0]
    try:
        vp.wait_for_load_state("domcontentloaded", timeout=30000)
    except Exception:
        pass
    time.sleep(5)
    return vp


def _open_video_page(page, folder_url: str, fname: str):
    srv_path = get_server_relative_path(folder_url, fname)
    if not srv_path:
        return None

    origin, base_path = _site_origin_base(folder_url)
    encoded_path = quote(srv_path, safe="")
    video_url = (
        f"{origin}{base_path}/_layouts/15/stream.aspx"
        f"?id={encoded_path}"
        f"&referrer=StreamWebApp.Web"
        f"&referrerScenario=AddressBarCopied.view"
    )

    video_page = page.context.new_page()
    video_page.goto(video_url, timeout=60000, wait_until="domcontentloaded")
    time.sleep(5)

    if _has_video_element(video_page):
        return video_page

    click.echo("    [WARN] Constructed URL failed (no <video>), falling back to click from folder...")
    video_page.close()
    return _open_video_by_click(page, folder_url, fname)


def _load_report(report_path: Path) -> dict:
    if report_path.exists():
        return json.loads(report_path.read_text(encoding="utf-8"))
    return {"completed": [], "failed": [], "skipped": []}


def _save_report(report_path: Path, report: dict):
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


@click.command()
@click.option("--url", required=True, help="URL folder SharePoint chua video.")
@click.option("--output-dir", default=str(DEFAULT_OUTPUT_DIR),
              help=f"Thu muc luu video (default: {DEFAULT_OUTPUT_DIR}).")
@click.option("--cdp", default="http://127.0.0.1:9222", help="Chrome DevTools Protocol URL.")
@click.option("--resume", is_flag=True, default=False,
              help="Tiep tuc tu lan chay truoc (bo qua file da thanh cong).")
@click.option("--retry-failed", is_flag=True, default=False,
              help="Chi chay lai cac file da that bai lan truoc.")
@click.option("--no-debug", is_flag=True, default=False, help="Tat debug log [CDP:DBG].")
@click.option("--max-duration", default=None, type=float,
              help="Gioi han thoi gian capture moi video (giay). Mac dinh: full video.")
def main(url, output_dir, cdp, resume, retry_failed, no_debug, max_duration):
    sys.stdout = _Tee(sys.__stdout__, LOG_FILE)
    sys.stderr = _Tee(sys.__stderr__, LOG_FILE)

    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    click.echo(f"\n{'='*60}")
    click.echo(f"  Batch capture started at {run_ts}")
    click.echo(f"  Log: {LOG_FILE}")
    click.echo(f"{'='*60}\n")

    if no_debug:
        video_capture._CDP_DEBUG = False

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    report_path = REPORT_DIR / "batch_report.json"
    report = _load_report(report_path)
    report.setdefault("started_at", datetime.now().isoformat())
    report.setdefault("completed", [])
    report.setdefault("failed", [])
    report.setdefault("skipped", [])

    with sync_playwright() as pw:
        try:
            browser = pw.chromium.connect_over_cdp(cdp)
        except Exception as e:
            click.echo(f"Khong ket noi duoc Chrome: {e}", err=True)
            click.echo("Mo Chrome voi remote debugging:")
            click.echo("  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome \\")
            click.echo("    --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-crawl")
            raise SystemExit(1)

        context = browser.contexts[0]
        page = context.pages[0]
        click.echo(f"Connected. {len(context.pages)} tab(s) open.\n")

        click.echo("--- Quet folder de tim file MP4 ---")
        all_mp4 = _discover_all_mp4(page, url)
        click.echo(f"Tim thay {len(all_mp4)} file MP4 trong folder.\n")

        if not all_mp4:
            click.echo("Khong co file MP4 nao.")
            return

        if retry_failed:
            failed_names = {r["name"] for r in report.get("failed", [])}
            all_mp4 = [f for f in all_mp4 if f["name"] in failed_names]
            report["failed"] = [r for r in report["failed"]
                                if r["name"] not in {f["name"] for f in all_mp4}]
            click.echo(f"Retry mode: {len(all_mp4)} file that bai can chay lai.\n")

        completed_names = {r["name"] for r in report.get("completed", [])}
        failed_names = {r["name"] for r in report.get("failed", [])}
        queue = []
        skipped = 0
        skip_failed = 0
        retry_invalid = 0
        for f in all_mp4:
            name = f["name"]
            safe = slugify(Path(name).stem) + ".mp4"
            out = output_dir / safe

            if out.exists() and _is_valid_video(out):
                if name not in completed_names:
                    report["completed"].append({
                        "name": name, "output": str(out),
                        "size_mb": round(out.stat().st_size / (1024 * 1024), 1),
                        "at": datetime.now().isoformat(),
                    })
                skipped += 1
                continue

            if name in failed_names and not retry_failed:
                skip_failed += 1
                continue

            if out.exists() and not _is_valid_video(out):
                size_kb = out.stat().st_size / 1024
                click.echo(f"  [RETRY] {name} (file {size_kb:.0f}KB khong hop le, xoa va chay lai)")
                out.unlink()
                report["completed"] = [r for r in report["completed"] if r["name"] != name]
                report["failed"] = [r for r in report["failed"] if r["name"] != name]
                retry_invalid += 1

            queue.append({"item": f, "output": out})

        if skipped:
            click.echo(f"  Bo qua {skipped} file da co video hop le.")
        if skip_failed:
            click.echo(f"  Bo qua {skip_failed} file da that bai truoc do (dung --retry-failed de thu lai).")
        if retry_invalid:
            click.echo(f"  Chay lai {retry_invalid} file co output khong hop le.")

        if not queue:
            click.echo("Khong con file nao can xu ly.")
            _save_report(report_path, report)
            _print_summary(report, report_path)
            return

        click.echo(f"\nSe xu ly {len(queue)} video.\n")
        click.echo("=" * 70)

        for idx, entry in enumerate(queue):
            fname = entry["item"]["name"]
            out = entry["output"]
            progress = f"[{idx + 1}/{len(queue)}]"

            click.echo(f"\n{progress} {fname}")
            click.echo(f"  Output: {out}")

            video_page = None
            t0 = time.time()
            try:
                video_page = _open_video_page(page, url, fname)
                if not video_page:
                    raise RuntimeError("Khong tao duoc video page URL")

                success = video_capture.capture_video_from_page(
                    video_page=video_page,
                    output_path=out,
                    strategy="auto",
                    max_duration=max_duration,
                )

                elapsed = time.time() - t0
                if success and out.exists() and out.stat().st_size > 100_000:
                    size_mb = out.stat().st_size / (1024 * 1024)
                    click.echo(f"  OK ({size_mb:.1f}MB, {elapsed:.0f}s)")
                    report["completed"].append({
                        "name": fname, "output": str(out),
                        "size_mb": round(size_mb, 1), "elapsed_s": round(elapsed),
                        "at": datetime.now().isoformat(),
                    })
                else:
                    click.echo(f"  FAILED (capture returned false, {elapsed:.0f}s)")
                    report["failed"].append({
                        "name": fname, "error": "capture returned false",
                        "elapsed_s": round(elapsed),
                        "at": datetime.now().isoformat(),
                    })

            except Exception as e:
                elapsed = time.time() - t0
                click.echo(f"  FAILED: {e} ({elapsed:.0f}s)")
                report["failed"].append({
                    "name": fname, "error": str(e)[:200],
                    "elapsed_s": round(elapsed),
                    "at": datetime.now().isoformat(),
                })

            finally:
                if video_page:
                    try:
                        video_page.close()
                    except Exception:
                        pass
                _save_report(report_path, report)

        _print_summary(report, report_path)


def _print_summary(report: dict, report_path: Path):
    click.echo("\n" + "=" * 70)
    click.echo("HOAN TAT")
    click.echo(f"  Thanh cong: {len(report['completed'])}")
    click.echo(f"  That bai:   {len(report['failed'])}")
    if report["failed"]:
        click.echo("\n  File that bai:")
        for r in report["failed"]:
            click.echo(f"    - {r['name']}: {r.get('error', '?')}")
        click.echo(f"\n  Chay lai file that bai:")
        click.echo(f"    uv run python batch_capture.py --url '...' --retry-failed")
    click.echo(f"\n  Report: {report_path}")
    click.echo(f"  Log: {LOG_FILE}")


if __name__ == "__main__":
    main()
