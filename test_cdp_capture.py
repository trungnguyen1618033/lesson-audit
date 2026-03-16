"""
test_cdp_capture.py -- Test CDP video packet capture tren SharePoint.

SETUP:
  1. uv sync
  2. uv run playwright install chromium
  3. brew install ffmpeg
  4. Mo Chrome voi remote debugging (tat Chrome thuong truoc):
       Xem README ben duoi
  5. Dang nhap SharePoint trong Chrome do.

CACH DUNG:
  # Mode 1: URL truc tiep den trang video
  uv run python test_cdp_capture.py --url "https://...stream.aspx?id=..."

  # Mode 2: URL folder + ten file cu the (script tu tim va mo video)
  uv run python test_cdp_capture.py \
    --url "https://tenant.sharepoint.com/.../AllItems.aspx?id=...Recordings..." \
    --target "Ten file video.mp4"

  # Gioi han 30s de test nhanh:
  uv run python test_cdp_capture.py --url "..." --target "..." --max-duration 30

  # Chi test CDP capture:
  uv run python test_cdp_capture.py --url "..." --target "..." --strategy cdp_capture

  # Dung tab hien tai (da mo san trang video):
  uv run python test_cdp_capture.py --current-tab --max-duration 30
"""
import time
from pathlib import Path
from urllib.parse import quote

import click
from playwright.sync_api import Page, sync_playwright

import video_capture


def _open_video_from_folder(page: Page, folder_url: str, target_name: str) -> Page:
    """Mo folder SharePoint, tim file target, mo trang Stream player. Tra ve page moi."""
    from crawler import (
        get_folder_items,
        get_server_relative_path,
        _site_origin_base,
        wait_for_sharepoint,
    )

    click.echo(f"  Mo folder: {folder_url[:80]}...")
    page.goto(folder_url, timeout=60000, wait_until="domcontentloaded")
    wait_for_sharepoint(page)

    click.echo("  Scroll de load het danh sach file...")
    for _ in range(5):
        page.mouse.wheel(0, 2000)
        time.sleep(1)

    items = get_folder_items(page)
    mp4_files = [i for i in items if i["type"] == "file" and i["name"].lower().endswith(".mp4")]
    click.echo(f"  Tim thay {len(mp4_files)} file MP4 trong folder.")

    if not mp4_files:
        raise click.ClickException("Khong tim thay file MP4 nao trong folder.")

    matched = [f for f in mp4_files if f["name"] == target_name]
    if not matched:
        click.echo(f"\n  Danh sach MP4 trong folder:")
        for f in mp4_files:
            click.echo(f"    - {f['name']}")
        raise click.ClickException(f"Khong tim thay file: {target_name!r}")

    fname = matched[0]["name"]
    click.echo(f"  Tim thay: {fname}")

    srv_path = get_server_relative_path(folder_url, fname)
    if not srv_path:
        raise click.ClickException(f"Khong lay duoc server relative path cho: {fname}")

    origin, base_path = _site_origin_base(folder_url)
    encoded_path = quote(srv_path, safe="")
    video_url = (
        f"{origin}{base_path}/_layouts/15/stream.aspx"
        f"?id={encoded_path}"
        f"&referrer=StreamWebApp.Web"
        f"&referrerScenario=AddressBarCopied.view"
    )
    click.echo(f"  Video URL: {video_url[:100]}...")

    video_page = page.context.new_page()
    video_page.goto(video_url, timeout=60000, wait_until="domcontentloaded")
    click.echo("  Doi trang video load...")
    time.sleep(5)

    return video_page


@click.command()
@click.option("--url", default=None, help="URL trang video hoac folder SharePoint.")
@click.option("--target", default=None, help="Ten file MP4 cu the trong folder (dung voi --url folder).")
@click.option("--current-tab", is_flag=True, default=False,
              help="Dung tab hien tai thay vi mo URL moi.")
@click.option("--strategy", default="auto",
              type=click.Choice(["auto", "intercept", "cdp_capture", "record"]),
              help="Chien luoc capture.")
@click.option("--max-duration", default=None, type=float,
              help="Gioi han thoi gian capture (giay).")
@click.option("--output", default=None, help="Duong dan file output (mac dinh: tu dong tu ten file).")
@click.option("--cdp", default="http://127.0.0.1:9222", help="Chrome DevTools Protocol URL.")
@click.option("--no-debug", is_flag=True, default=False, help="Tat debug log [CDP:DBG].")
def main(url, target, current_tab, strategy, max_duration, output, cdp, no_debug):
    if no_debug:
        video_capture._CDP_DEBUG = False

    if not url and not current_tab:
        click.echo("Phai truyen --url hoac --current-tab")
        click.echo("\nVi du:")
        click.echo('  uv run python test_cdp_capture.py --url "https://...folder..." --target "video.mp4" --max-duration 30')
        click.echo('  uv run python test_cdp_capture.py --current-tab --max-duration 30')
        raise SystemExit(1)

    if output:
        output_path = Path(output)
    elif target:
        from crawler import slugify
        safe_name = slugify(Path(target).stem)
        output_path = Path(f"captures/_test_video/{safe_name}.mp4")
    else:
        output_path = Path("captures/_test_video/test_output.mp4")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    click.echo(f"Strategy:     {strategy}")
    click.echo(f"Max duration: {max_duration or 'full video'}")
    click.echo(f"Output:       {output_path}")
    click.echo(f"CDP debug:    {'OFF' if no_debug else 'ON'}")
    if target:
        click.echo(f"Target file:  {target}")
    click.echo()

    with sync_playwright() as pw:
        try:
            browser = pw.chromium.connect_over_cdp(cdp)
        except Exception as e:
            click.echo(f"\nKhong ket noi duoc Chrome: {e}", err=True)
            click.echo("\nMo Chrome voi remote debugging:")
            click.echo('  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome \\')
            click.echo("    --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-crawl")
            raise SystemExit(1)

        context = browser.contexts[0]
        click.echo(f"Connected. {len(context.pages)} tab(s) open.")

        video_page = None
        success = False

        try:
            if current_tab:
                page = context.pages[0]
                click.echo(f"Using current tab: {page.url[:80]}")
                click.echo()
                success = video_capture.capture_video_from_page(
                    video_page=page,
                    output_path=output_path,
                    strategy=strategy,
                    max_duration=max_duration,
                )

            elif target:
                page = context.pages[0]
                click.echo(f"Mode: folder + target file\n")
                video_page = _open_video_from_folder(page, url, target)
                click.echo()
                success = video_capture.capture_video_from_page(
                    video_page=video_page,
                    output_path=output_path,
                    strategy=strategy,
                    max_duration=max_duration,
                )

            else:
                page = context.pages[0]
                click.echo(f"Opening: {url[:80]}...")
                click.echo()
                success = video_capture.capture_video(
                    page=page,
                    video_page_url=url,
                    output_path=output_path,
                    strategy=strategy,
                    max_duration=max_duration,
                )

        finally:
            if video_page:
                try:
                    video_page.close()
                except Exception:
                    pass

        click.echo()
        if success and output_path.exists():
            size_mb = output_path.stat().st_size / (1024 * 1024)
            click.echo(f"OK: {output_path} ({size_mb:.1f}MB)")
        else:
            click.echo("FAILED: Video capture khong thanh cong.")
            click.echo("Xem log [CDP:DBG] phia tren de debug.")
            raise SystemExit(1)


if __name__ == "__main__":
    main()
