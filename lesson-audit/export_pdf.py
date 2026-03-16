"""
export_pdf.py — Tạo file PDF từ slide PNG cho mỗi bài giảng trong thư mục captures.

Cách dùng:
  uv run python export_pdf.py                    # Tạo PDF cho tất cả
  uv run python export_pdf.py --output pdf_out   # Chỉ định thư mục output
  uv run python export_pdf.py --flat             # Lưu tất cả PDF vào 1 thư mục phẳng
"""

import sys
from pathlib import Path
from PIL import Image
import click


CAPTURES_DIR = Path("captures")


def find_presentations(captures_dir: Path) -> list[Path]:
    """Find all presentation folders (folders containing slide_001.png)."""
    results = []
    for slide1 in sorted(captures_dir.rglob("slide_001.png")):
        results.append(slide1.parent)
    return results


def slides_to_pdf(folder: Path, output_path: Path) -> int:
    """Convert slide_*.png in folder to a single PDF. Returns slide count."""
    slide_files = sorted(folder.glob("slide_*.png"))
    if not slide_files:
        return 0

    images = []
    for sf in slide_files:
        img = Image.open(sf).convert("RGB")
        images.append(img)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    images[0].save(output_path, "PDF", save_all=True, append_images=images[1:])
    return len(images)


@click.command()
@click.option("--output", "output_dir", default="pdf_output", show_default=True,
              help="Thư mục lưu PDF.")
@click.option("--flat", is_flag=True, default=False,
              help="Lưu tất cả PDF vào 1 thư mục phẳng (không giữ cấu trúc thư mục).")
def main(output_dir: str, flat: bool):
    """Tạo file PDF từ slide PNG cho mỗi bài giảng."""
    out = Path(output_dir)
    presentations = find_presentations(CAPTURES_DIR)

    if not presentations:
        click.echo("Không tìm thấy bài giảng nào trong captures/")
        return

    click.echo(f"Tìm thấy {len(presentations)} bài giảng. Đang tạo PDF...\n")

    created = 0
    for folder in presentations:
        rel = folder.relative_to(CAPTURES_DIR)
        name = folder.name

        if flat:
            pdf_name = str(rel).replace("/", " - ") + ".pdf"
            pdf_path = out / pdf_name
        else:
            pdf_path = out / rel / f"{name}.pdf"

        if pdf_path.exists():
            slide_count = len(list(folder.glob("slide_*.png")))
            click.echo(f"  [SKIP] {rel} ({slide_count} slides, PDF exists)")
            created += 1
            continue

        count = slides_to_pdf(folder, pdf_path)
        if count > 0:
            click.echo(f"  [OK] {rel} → {count} slides → {pdf_path.name}")
            created += 1

    click.echo(f"\nHoàn thành! {created}/{len(presentations)} PDF trong: {out}/")


if __name__ == "__main__":
    main()
