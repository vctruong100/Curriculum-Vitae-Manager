"""
Build-time icon generator for CV Research Experience Manager.

Called automatically by build scripts before PyInstaller runs.
Priority order:
  1. build/assets/feather.ico  — user-supplied custom icon (copied as-is)
  2. Fallback                  — programmatic feather pen icon via Pillow

Output: build/assets/app.ico  (multi-size ICO used by PyInstaller and the GUI)
Requires Pillow (pip install Pillow) for the fallback path.
No network requests — purely local file generation.
"""

import sys
import os
import math
import shutil
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

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
    max_size = max(ICO_SIZES)
    img = img.resize((max_size, max_size), Image.LANCZOS)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(
        str(output_path),
        format="ICO",
        sizes=[(s, s) for s in ICO_SIZES],
    )
    return output_path


def _lerp_profile(profile, t):
    """Linearly interpolate left/right widths from the profile at parameter t."""
    for i in range(len(profile) - 1):
        t0, lw0, rw0 = profile[i]
        t1, lw1, rw1 = profile[i + 1]
        if t0 <= t <= t1:
            f = (t - t0) / (t1 - t0) if t1 != t0 else 0.0
            return lw0 + f * (lw1 - lw0), rw0 + f * (rw1 - rw0)
    return 0.0, 0.0


def generate_feather_pen_icon(output_path: Path) -> Path:
    """Generate a classic feather pen icon (deep navy background, white quill, gold nib)."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("Pillow is not installed. Install with: pip install Pillow")
        print("Skipping icon generation.")
        return None

    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # ── Background: rounded square, deep navy ──
    pad = 6
    bg_color = (30, 58, 95, 255)
    try:
        draw.rounded_rectangle([pad, pad, size - pad, size - pad], radius=36, fill=bg_color)
    except AttributeError:
        draw.rectangle([pad, pad, size - pad, size - pad], fill=bg_color)

    # ── Feather axis from nib (lower-left) to tip (upper-right) ──
    ax0, ay0 = 62, 194
    ax1, ay1 = 194, 62
    dx = ax1 - ax0
    dy = ay1 - ay0
    length = math.hypot(dx, dy)
    ux, uy = dx / length, dy / length
    px, py = -uy, ux

    # Width profile: (t, left_width, right_width).  t=0 is nib, t=1 is feather tip.
    profile = [
        (0.00,  1,  1),
        (0.12,  2,  2),
        (0.22,  3,  3),
        (0.32, 14, 11),
        (0.42, 24, 19),
        (0.52, 32, 25),
        (0.62, 37, 29),
        (0.72, 38, 30),
        (0.82, 33, 26),
        (0.90, 22, 17),
        (0.96,  8,  6),
        (1.00,  1,  1),
    ]

    left_pts, right_pts = [], []
    for t, lw, rw in profile:
        cx = ax0 + t * dx
        cy = ay0 + t * dy
        left_pts.append((cx + lw * px, cy + lw * py))
        right_pts.append((cx - rw * px, cy - rw * py))

    feather_poly = left_pts + list(reversed(right_pts))
    draw.polygon(feather_poly, fill=(248, 246, 240, 255))

    # ── Shaft / rachis ──
    s0 = (ax0 + 0.15 * dx, ay0 + 0.15 * dy)
    s1 = (ax0 + 0.97 * dx, ay0 + 0.97 * dy)
    draw.line([s0, s1], fill=(185, 175, 158, 255), width=2)

    # ── Feather barb lines (subtle diagonal texture) ──
    barb_color = (210, 200, 185, 160)
    for t in (0.40, 0.50, 0.60, 0.70, 0.80, 0.88):
        cx = ax0 + t * dx
        cy = ay0 + t * dy
        lw, rw = _lerp_profile(profile, t)
        left_end = (cx + lw * 0.85 * px, cy + lw * 0.85 * py)
        right_end = (cx - rw * 0.85 * px, cy - rw * 0.85 * py)
        draw.line([(cx, cy), left_end], fill=barb_color, width=1)
        draw.line([(cx, cy), right_end], fill=barb_color, width=1)

    # ── Gold nib ──
    nib_t = 0.14
    ncx = ax0 + nib_t * dx
    ncy = ay0 + nib_t * dy
    nib_poly = [
        (ax0, ay0),
        (ncx + 4 * px, ncy + 4 * py),
        (ncx - 4 * px, ncy - 4 * py),
    ]
    draw.polygon(nib_poly, fill=(205, 165, 50, 255))

    # Nib slit
    slit_end = (ax0 + 0.08 * dx, ay0 + 0.08 * dy)
    draw.line([(ax0, ay0), slit_end], fill=bg_color, width=1)

    # ── Save as multi-size ICO ──
    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(
        str(output_path),
        format="ICO",
        sizes=[(s, s) for s in ICO_SIZES],
    )
    return output_path


def generate_icon(output_path: Path, logo_png: Path = None) -> Path:
    """Generate .ico.  Uses Logo.png if provided, otherwise draws a feather pen."""
    if logo_png is not None and logo_png.exists():
        print(f"Converting {logo_png} -> {output_path}")
        return generate_icon_from_png(logo_png, output_path)
    else:
        print("Generating feather pen icon.")
        return generate_feather_pen_icon(output_path)


def main():
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    build_dir = Path(__file__).parent.resolve()
    assets_dir = build_dir / "assets"
    icon_path = assets_dir / "app.ico"
    feather_ico = assets_dir / "feather.ico"

    # Priority 1: user-supplied feather.ico
    if feather_ico.exists():
        logger.info("Found custom icon at %s — copying to %s", feather_ico, icon_path)
        assets_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(feather_ico), str(icon_path))
        logger.info(
            "Icon ready at %s (%d bytes)",
            icon_path,
            icon_path.stat().st_size,
        )
        sys.exit(0)

    # Priority 2: generate feather pen icon via Pillow
    logger.info("No feather.ico found — generating feather pen icon via Pillow.")
    result = generate_feather_pen_icon(icon_path)

    if result is not None:
        logger.info(
            "Icon generated at %s (%d bytes)",
            result,
            result.stat().st_size,
        )
    else:
        logger.error("Icon generation failed.")
        sys.exit(1)


if __name__ == "__main__":
    main()
