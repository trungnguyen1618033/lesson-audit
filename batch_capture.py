import json
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import click
from playwright.sync_api import sync_playwright

import video_capture
from crawler import (
    get_folder_items,
    get_server_relative_path,
    _site_origin_base,
    wait_for_sharepoint,
    slugify,
)

REPORT_DIR = Path("captures/_reports")
DEFAULT_OUTPUT_DIR = Path("captures/recordings")


def _discover_all_mp4(page, folder_url: str, max_scroll: int = 30) -> list[dict]:
    page.goto(folder_url, timeout=60000, wait_until="domcontentloaded")
    wait_for_sharepoint(page)

    all_mp4: dict[str, dict] = {}
    stable_count = 0

    for i in range(max_scroll):
        page.mouse.wheel(0, 3000)
        time.sleep(1.5)
        items = get_folder_items(page)
        mp4s = [it for it in items if it["type"] == "file" and it["name"].lower().endswith(".mp4")]
        for f in mp4s:
            all_mp4[f["name"]] = f

        if len(all_mp4) == len(mp4s) and i > 3:
            stable_count += 1
            if stable_count >= 3:
                break
        else:
            stable_count = 0

    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(1)
    items = get_folder_items(page)
    for f in items:
        if f["type"] == "file" and f["name"].lower().endswith(".mp4"):
            all_mp4[f["name"]] = f

    return sorted(all_mp4.values(), key=lambda x: x["name"])


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
    return video_page


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
    if no_debug:
        video_capture._CDP_DEBUG = False

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    report_path = REPORT_DIR / "batch_report.json"
    report = _load_report(report_path) if (resume or retry_failed) else {
        "completed": [], "failed": [], "skipped": [],
        "started_at": datetime.now().isoformat(),
    }
    completed_names = {r["name"] for r in report.get("completed", [])}

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

        click.echo("Dang quet folder de tim tat ca file MP4...")
        all_mp4 = _discover_all_mp4(page, url)
        click.echo(f"Tim thay {len(all_mp4)} file MP4 trong folder.\n")

        if not all_mp4:
            click.echo("Khong co file MP4 nao.")
            raise SystemExit(0)

        if retry_failed:
            failed_names = {r["name"] for r in report.get("failed", [])}
            all_mp4 = [f for f in all_mp4 if f["name"] in failed_names]
            report["failed"] = [r for r in report["failed"] if r["name"] not in {f["name"] for f in all_mp4}]
            click.echo(f"Retry mode: {len(all_mp4)} file that bai can chay lai.\n")

        queue = []
        for f in all_mp4:
            name = f["name"]
            safe = slugify(Path(name).stem) + ".mp4"
            out = output_dir / safe

            if name in completed_names and out.exists():
                click.echo(f"  SKIP (da xong): {name}")
                continue
            queue.append({"item": f, "output": out})

        click.echo(f"\nSe xu ly {len(queue)} / {len(all_mp4)} video.\n")
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


if __name__ == "__main__":
    main()
