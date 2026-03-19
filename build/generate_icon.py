"""
Build-time icon generator for CV Research Experience Manager.

Usage:
    python build/generate_icon.py

Converts build/assets/Logo.png into build/assets/app.ico (multi-size ICO).
Falls back to generating a simple 'CV' placeholder if Logo.png is missing.
Requires Pillow (pip install Pillow).
No network requests — purely local file generation.
"""

import sys
import os
from pathlib import Path


ICO_SIZES = [16, 32, 48, 64, 128, 256]


def generate_icon_from_png(png_path: Path, output_path: Path) -> Path:
    """Convert an existing PNG logo into a multi-size .ico file."""
    try:
        from PIL import Image
    except ImportError:
        print("Pillow is not installed. Install with: pip install Pillow")
        print("Skipping icon generation.")
        return None

    img = Image.open(png_path).convert("RGBA")
    images = []
    for size in ICO_SIZES:
        resized = img.resize((size, size), Image.LANCZOS)
        images.append(resized)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    images[0].save(
        str(output_path),
        format="ICO",
        sizes=[(s, s) for s in ICO_SIZES],
        append_images=images[1:],
    )
    return output_path


def generate_icon_fallback(output_path: Path) -> Path:
    """Generate a simple 'CV' placeholder icon when Logo.png is missing."""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("Pillow is not installed. Install with: pip install Pillow")
        print("Skipping icon generation.")
        return None

    images = []
    for size in ICO_SIZES:
        img = Image.new("RGBA", (size, size), (37, 99, 235, 255))
        draw = ImageDraw.Draw(img)

        font_size = int(size * 0.45)
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except (IOError, OSError):
            try:
                font = ImageFont.truetype("DejaVuSans-Bold.ttf", font_size)
            except (IOError, OSError):
                font = ImageFont.load_default()

        text = "CV"
        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x = (size - text_w) // 2
        y = (size - text_h) // 2 - bbox[1]
        draw.text((x, y), text, fill="white", font=font)
        images.append(img)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    images[0].save(
        str(output_path),
        format="ICO",
        sizes=[(s, s) for s in ICO_SIZES],
        append_images=images[1:],
    )
    return output_path


def generate_icon(output_path: Path, logo_png: Path = None) -> Path:
    """Generate .ico, preferring Logo.png conversion over fallback."""
    if logo_png is not None and logo_png.exists():
        print(f"Converting {logo_png} -> {output_path}")
        return generate_icon_from_png(logo_png, output_path)
    else:
        print("Logo.png not found — generating fallback icon.")
        return generate_icon_fallback(output_path)


def main():
    build_dir = Path(__file__).parent.resolve()
    assets_dir = build_dir / "assets"
    logo_png = assets_dir / "Logo.png"
    icon_path = assets_dir / "app.ico"

    if icon_path.exists():
        print(f"Icon already exists at {icon_path} — skipping generation.")
        return

    result = generate_icon(icon_path, logo_png)
    if result is not None:
        print(f"Icon generated at {result}")
    else:
        print("Icon generation failed.")
        sys.exit(1)


if __name__ == "__main__":
    main()
