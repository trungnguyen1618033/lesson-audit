"""
main.py - CLI entry point for the Auto Capture PowerPoint Slides tool.

Commands
--------
  setup-profile       Lưu vùng slide + nút next (chạy 1 lần)
  update-present-btn  Cập nhật vị trí nút Trình bày
  list-profiles       Xem danh sách profiles
  capture             Tự động chụp từng slide và xuất PPTX
  assemble            Ráp ảnh đã có thành PPTX

Ví dụ
-----
  python main.py setup-profile --name pdf-viewer
  python main.py capture --name "Bai_giang" --profile pptx-preview
  python main.py capture --name "1._NLS/session/file" --profile pptx-preview --skip-present --auto-close
"""

import time
from pathlib import Path

import click
import pyautogui

from assembler import assemble_from_dir, assemble_pptx
from capturer import get_monitor_under_mouse, run_capture_session
from navigator import press_next, set_nav_key, set_next_button_position
from profiles import get_profile, list_profiles, setup_profile

CAPTURES_DIR = Path("captures")


@click.group()
def cli():
    """Auto Capture PowerPoint Slides - macOS screen capture tool."""


# ── setup-profile ──────────────────────────────────────────────────────────────

@cli.command("setup-profile")
@click.option("--name", required=True,
              prompt="Tên profile (vd: pdf-viewer, pptx-preview)",
              help="Tên định danh cho profile này.")
def cmd_setup_profile(name):
    """Lưu vùng slide + vị trí nút next cho một loại viewer. Chạy 1 lần, dùng mãi."""
    setup_profile(name)


# ── update-present-btn ────────────────────────────────────────────────────────

@cli.command("update-present-btn")
@click.option("--name", required=True,
              help="Tên profile cần cập nhật (vd: pptx-preview).")
def cmd_update_present_btn(name):
    """Cập nhật vị trí nút Trình bày (không đổi các cài đặt khác)."""
    from profiles import load_profiles, save_profiles
    p = get_profile(name)
    if p is None:
        click.echo(f"Profile '{name}' không tồn tại.", err=True)
        raise SystemExit(1)

    if p.get("present_btn"):
        click.echo(f"Hiện tại: x={p['present_btn']['x']}, y={p['present_btn']['y']}")

    click.echo("\nMở PowerPoint Online trong Chrome, CHƯA bấm gì.")
    input("Bấm Enter khi sẵn sàng (có 8 giây để di chuột đến nút Trình bày): ")

    for i in range(8, 0, -1):
        pos = pyautogui.position()
        click.echo(f"\r  [{i}s] x={pos.x}, y={pos.y}   ", nl=False)
        time.sleep(1)
    pos = pyautogui.position()
    click.echo(f"\n  [OK] Vị trí nút Trình bày: x={pos.x}, y={pos.y}")

    profiles = load_profiles()
    profiles[name]["present_btn"] = {"x": pos.x, "y": pos.y}
    save_profiles(profiles)
    click.echo(f"Profile '{name}' đã cập nhật!")


# ── list-profiles ─────────────────────────────────────────────────────────────

@cli.command("list-profiles")
def cmd_list_profiles():
    """Xem danh sách profiles đã lưu."""
    from profiles import load_profiles
    names = list_profiles()
    if not names:
        click.echo("Chưa có profile nào. Chạy: python main.py setup-profile --name <tên>")
        return
    profiles = load_profiles()
    for n in names:
        p = profiles[n]
        r = p.get("region")
        region_info = f"{r['width']}x{r['height']} @ ({r['left']},{r['top']})" if r else "full screen"
        nav_info = (f"click ({p['nav']['x']},{p['nav']['y']})"
                    if p.get("nav") else f"key [{p.get('key', 'right')}]")
        present = f"  present_btn=({p['present_btn']['x']},{p['present_btn']['y']})" \
                  if p.get("present_btn") else ""
        click.echo(f"  {n:20s}  region={region_info}  nav={nav_info}{present}")


# ── capture ───────────────────────────────────────────────────────────────────

@cli.command("capture")
@click.option("--name", required=True,
              prompt="Tên presentation (dùng làm tên thư mục và file PPTX)",
              help="Tên thư mục output. Có thể dùng dấu / để tạo subfolder.")
@click.option("--profile", default=None,
              help="Tên profile đã lưu (vd: pdf-viewer, pptx-preview).")
@click.option("--delay", default=0.0, show_default=True,
              help="Giây chờ thêm sau mỗi lần nhấn Next.")
@click.option("--max-slides", default=200, show_default=True,
              help="Số slide tối đa.")
@click.option("--diff-threshold", default=1.0, show_default=True,
              help="% pixel diff để coi là slide mới.")
@click.option("--same-count", default=10, show_default=True,
              help="Số lần liên tiếp không đổi thì dừng.")
@click.option("--total", default=None, type=int,
              help="Dừng chính xác sau N slides (bỏ qua auto-detect).")
@click.option("--no-pptx", is_flag=True, default=False,
              help="Chỉ lưu ảnh, không tạo PPTX.")
@click.option("--auto-close", is_flag=True, default=False,
              help="Tự đóng tab trình duyệt sau khi xong.")
@click.option("--nav-x", default=None, type=int, help="Tọa độ x nút Next (ghi đè profile).")
@click.option("--nav-y", default=None, type=int, help="Tọa độ y nút Next (ghi đè profile).")
@click.option("--nav-key", default=None, help="Phím điều hướng (vd: right, pagedown).")
@click.option("--skip-present", is_flag=True, default=False,
              help="Bỏ qua bước click Trình bày (crawler đã xử lý).")
@click.option("--force", is_flag=True, default=False,
              help="Ghi đè folder đã có mà không hỏi.")
@click.option("--region", default=None,
              help="Vùng chụp màn hình 'left,top,width,height' (ghi đè profile region). "
                   "Dùng khi crawler tự detect.")
def cmd_capture(name, profile, delay, max_slides, diff_threshold, same_count,
                total, no_pptx, auto_close, nav_x, nav_y, nav_key, skip_present, force, region):
    """
    Tự động chụp từng slide và xuất file PPTX.

    Khi chạy tay: có countdown 5s để di chuột sang Chrome.
    Khi gọi từ crawler (--skip-present): bỏ qua countdown và click Trình bày.
    """
    output_dir = CAPTURES_DIR / name

    # Overwrite check — skip when called from crawler
    if not skip_present and not force:
        if output_dir.exists() and any(output_dir.glob("slide_*.png")):
            if not click.confirm(f"Folder '{output_dir}' đã có ảnh. Ghi đè?", default=False):
                raise SystemExit(0)

    # ── Load profile ──────────────────────────────────────────────────────────
    p = get_profile(profile) if profile else None
    if profile and p is None:
        click.echo(f"Profile '{profile}' không tồn tại.", err=True)
        raise SystemExit(1)

    # ── Nav button / key ──────────────────────────────────────────────────────
    if nav_x is not None and nav_y is not None:
        set_next_button_position(nav_x, nav_y)
    elif p and p.get("nav"):
        set_next_button_position(p["nav"]["x"], p["nav"]["y"])

    if nav_key:
        set_nav_key(nav_key)
    elif p:
        set_nav_key(p.get("key", "right"))

    # ── Determine capture region ──────────────────────────────────────────────
    # Parse --region "L,T,W,H" if provided (overrides everything)
    if region:
        try:
            l, t, w, h = [int(x) for x in region.split(",")]
            region = {"left": l, "top": t, "width": w, "height": h}
            click.echo(f"\nAuto-detected region: {region['width']}x{region['height']}"
                       f" @ ({region['left']},{region['top']})")
        except Exception:
            click.echo(f"⚠ Invalid --region '{region}', ignoring.", err=True)
            region = None

    if region:
        pass  # already set above
    elif p and p.get("region"):
        region = p["region"]
        click.echo(f"\nProfile crop: {region['width']}x{region['height']} @ ({region['left']},{region['top']})")
    elif skip_present:
        # Mouse was moved to Chrome by crawler — capture that monitor
        region = get_monitor_under_mouse()
        click.echo(f"\nCapture region: {region['width']}x{region['height']} px")
    else:
        click.echo("\nBước 1 — Di chuột sang màn hình Chrome trong 5 giây...")
        for i in range(5, 0, -1):
            pos = pyautogui.position()
            click.echo(f"\r  [{i}s] Chuột đang ở: x={pos.x:5d}, y={pos.y:5d}   ", nl=False)
            time.sleep(1)
        region = get_monitor_under_mouse()
        click.echo(f"\nChụp toàn màn hình: {region['width']}x{region['height']} px")

    # ── Start presentation ────────────────────────────────────────────────────
    in_presentation_mode = False

    if skip_present:
        in_presentation_mode = True  # crawler handled it
        # Re-activate Chrome so keyboard events (right arrow) go to the slideshow.
        # The crawler may have switched back to the terminal to launch main.py.
        try:
            import subprocess
            subprocess.run(
                ["osascript", "-e", 'tell application "Google Chrome" to activate'],
                capture_output=True, timeout=3,
            )
        except Exception:
            pass
        time.sleep(1.5)  # let slideshow finish starting before first capture
    elif p and p.get("present_btn"):
        btn = p["present_btn"]
        click.echo(f"\nBước 2 — Click nút Trình bày ({btn['x']}, {btn['y']})...")
        time.sleep(0.5)
        pyautogui.click(btn["x"], btn["y"])
        time.sleep(2.5)
        click.echo("Đã vào trình chiếu.")
        in_presentation_mode = True
    else:
        click.echo("\nBước 2 — Vào chế độ Trình bày, rồi quay lại đây.")
        input("Nhấn Enter khi sẵn sàng: ")

    # ── Go to slide 1 (manual mode) ───────────────────────────────────────────
    if not in_presentation_mode:
        from navigator import click_slide_area, press_home
        click_slide_area(region)
        time.sleep(0.3)
        press_home()
        time.sleep(1)

    # ── Capture loop ──────────────────────────────────────────────────────────
    click.echo(f"\nCapturing into: {output_dir}")
    saved_paths = run_capture_session(
        region=region,
        output_dir=output_dir,
        navigate_fn=press_next,
        delay=delay,
        max_slides=total if total else max_slides,
        diff_threshold=diff_threshold,
        same_count_limit=same_count,
        exact_total=total,
    )

    if not saved_paths:
        click.echo("Không có slide nào được chụp.", err=True)
        raise SystemExit(1)

    # ── Close tab ─────────────────────────────────────────────────────────────
    if auto_close:
        time.sleep(0.3)
        pyautogui.hotkey("command", "w")
        click.echo("Tab đã đóng.")

    # ── Assemble PPTX ─────────────────────────────────────────────────────────
    if not no_pptx:
        pptx_name = Path(name).name  # last segment only (no slash duplication)
        pptx_path = output_dir / f"{pptx_name}.pptx"
        assemble_pptx(saved_paths, pptx_path)
        click.echo(f"\nHoàn thành! PPTX: {pptx_path}")
    else:
        click.echo(f"\nHoàn thành! Ảnh lưu tại: {output_dir}")

    # ── Wait for Enter only in interactive mode ───────────────────────────────
    if not skip_present:
        input("\nNhấn Enter để thoát...")


# ── assemble ──────────────────────────────────────────────────────────────────

@cli.command("assemble")
@click.option("--dir", "image_dir", required=True,
              help="Thư mục chứa slide_NNN.png.")
@click.option("--name", required=True,
              help="Tên file PPTX output (không cần .pptx).")
def cmd_assemble(image_dir, name):
    """Ráp ảnh trong thư mục thành file PPTX (không cần chụp lại)."""
    image_dir_path = Path(image_dir)
    if not image_dir_path.exists():
        click.echo(f"Không tìm thấy thư mục: {image_dir}", err=True)
        raise SystemExit(1)
    try:
        pptx_path = assemble_from_dir(image_dir_path, name)
        click.echo(f"Hoàn thành! PPTX: {pptx_path}")
    except FileNotFoundError as e:
        click.echo(str(e), err=True)
        raise SystemExit(1)


if __name__ == "__main__":
    cli()
