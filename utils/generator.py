"""
generator.py  (v2 — with Google Fonts + bold/italic + custom font support)
--------------------------------------------------------------------------
Core certificate generation engine using Pillow (PIL).

Each field dict now carries extra keys:
  - fontFamily : Google Font name OR "custom"    (e.g. "Roboto", "Montserrat")
  - fontPath   : resolved absolute path to the .ttf / .otf font file
  - bold       : bool – apply bold variant if available
  - italic     : bool – apply italic variant if available
"""

import os
import re
import urllib.request
import urllib.parse
from PIL import Image, ImageDraw, ImageFont


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sanitize_filename(name: str) -> str:
    """Remove characters illegal in Windows filenames."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def _load_font(font_path: str | None, size: int) -> ImageFont.FreeTypeFont:
    """
    Load a TrueType / OpenType font by file path.
    Falls back to Windows Arial → Calibri → Pillow default.
    """
    if font_path and os.path.isfile(font_path):
        try:
            return ImageFont.truetype(font_path, size)
        except Exception:
            pass  # Corrupt or wrong file — fall through

    # Common Windows system font fallbacks
    for candidate in [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/calibri.ttf",
        "C:/Windows/Fonts/times.ttf",
    ]:
        if os.path.isfile(candidate):
            return ImageFont.truetype(candidate, size)

    return ImageFont.load_default()


def _auto_shrink_font(
    draw: ImageDraw.ImageDraw,
    text: str,
    font_path: str | None,
    initial_size: int,
    max_width: int,
    min_size: int = 10,
) -> ImageFont.FreeTypeFont:
    """
    Iteratively reduce font size until the rendered text fits within max_width.
    Stops at min_size to prevent illegible output.
    """
    size = initial_size
    while size >= min_size:
        font = _load_font(font_path, size)
        # Check multiline width
        lines = text.split('\n')
        max_line_width = max(draw.textbbox((0, 0), line, font=font)[2] for line in lines)
        if max_line_width <= max_width:
            return font
        size -= 1
    return _load_font(font_path, min_size)


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert '#RRGGBB' or '#RGB' to an (R, G, B) int tuple."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) == 3:
        hex_color = "".join(c * 2 for c in hex_color)
    return tuple(int(hex_color[i: i + 2], 16) for i in (0, 2, 4))


def _resolve_font_path(
    font_path: str | None,
    bold: bool,
    italic: bool,
    fonts_dir: str,
) -> str | None:
    """
    Given a base font path and bold/italic flags, try to find the best
    matching variant file in the same directory.

    Variant naming conventions (Google Fonts uses these):
      Regular : FontName-Regular.ttf
      Bold    : FontName-Bold.ttf
      Italic  : FontName-Italic.ttf
      BoldIta : FontName-BoldItalic.ttf
    """
    if not font_path or not os.path.isfile(font_path):
        return font_path

    directory = os.path.dirname(font_path)
    basename  = os.path.basename(font_path)
    name_no_ext, ext = os.path.splitext(basename)

    # Strip known variant suffixes to get the base family name
    suffixes_to_strip = [
        "-Regular", "-Bold", "-Italic", "-BoldItalic",
        "-Light", "-Medium", "-SemiBold", "-ExtraBold", "-Black",
        "Regular", "Bold", "Italic",
    ]
    family_base = name_no_ext
    for s in suffixes_to_strip:
        if family_base.endswith(s):
            family_base = family_base[: -len(s)]
            break

    # Determine desired variant suffix
    if bold and italic:
        candidates = [f"{family_base}-BoldItalic{ext}", f"{family_base}-Bold{ext}"]
    elif bold:
        candidates = [f"{family_base}-Bold{ext}", f"{family_base}-SemiBold{ext}"]
    elif italic:
        candidates = [f"{family_base}-Italic{ext}"]
    else:
        candidates = [f"{family_base}-Regular{ext}", basename]

    for c in candidates:
        full = os.path.join(directory, c)
        if os.path.isfile(full):
            return full

    return font_path


def _download_google_font(family: str, bold: bool, italic: bool, fonts_dir: str) -> str | None:
    """
    Downloads a Google Font `.ttf` into `fonts_dir` if it doesn't already exist.
    """
    if not family:
        return None

    os.makedirs(fonts_dir, exist_ok=True)
    
    # Construct a safe filename based on variants
    suffix = ""
    if bold and italic: suffix = "-BoldItalic"
    elif bold: suffix = "-Bold"
    elif italic: suffix = "-Italic"
    else: suffix = "-Regular"
    
    safe_family = family.replace(" ", "")
    out_path = os.path.join(fonts_dir, f"{safe_family}{suffix}.ttf")
    if os.path.isfile(out_path):
        return out_path

    # Old wget user-agent tricks Google into serving raw .ttf files
    # We request the family and weight/style combination if needed.
    # Actually, standard query: family=Montserrat:wght@400;700
    # Or CSS1: family=Montserrat:400,700,400italic,700italic
    req_family = urllib.parse.quote(family)
    weight = "700" if bold else "400"
    style = "italic" if italic else ""
    # E.g. Roboto:700italic
    url = f"https://fonts.googleapis.com/css?family={req_family}:{weight}{style}"
    
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "wget/1.20"})
        css = urllib.request.urlopen(req, timeout=5).read().decode('utf-8')
        # Find url(https://...ttf)
        match = re.search(r"url\((https://[^)]+\.ttf)\)", css)
        if match:
            ttf_url = match.group(1)
            ttf_data = urllib.request.urlopen(ttf_url, timeout=10).read()
            with open(out_path, "wb") as f:
                f.write(ttf_data)
            return out_path
    except Exception as e:
        print(f"Failed to download Google Font {family}: {e}")

    # Fallback: original font path unchanged
    return None


# ---------------------------------------------------------------------------
# Main generator function
# ---------------------------------------------------------------------------

def generate_certificate(
    template_path: str,
    row: dict,
    fields: list[dict],
    output_dir: str,
    fonts_dir: str | None = None,
) -> str:
    """
    Generate a single certificate image for one participant row.

    Args:
        template_path : Path to the certificate template (JPEG/PNG).
        row           : Dict of {column_name: value} for one participant.
        fields        : List of field config dicts. Recognised keys:
                          column      – Excel column name
                          x, y        – Image-space anchor coordinates
                          fontSize    – integer point size
                          color       – hex colour string e.g. "#ffffff"
                          align       – "left" | "center" | "right"
                          fontPath    – absolute path to .ttf/.otf file
                          bold        – bool
                          italic      – bool
        output_dir    : Directory where {Name}.png will be saved.
        fonts_dir     : Directory containing cached Google Font / custom files.

    Returns:
        Absolute path of the saved PNG.
    """
    os.makedirs(output_dir, exist_ok=True)

    img = Image.open(template_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    img_width, img_height = img.size

    for field in fields:
        column    = field.get("column", "")
        text      = str(row.get(column, "")).strip()
        if not text or text.lower() == "nan":
            continue

        x          = float(field.get("x", 0))
        y          = float(field.get("y", 0))
        base_size  = int(field.get("fontSize", 30))
        color_hex  = field.get("color", "#000000")
        alignment  = field.get("align", "center")
        bold       = bool(field.get("bold", False))
        italic     = bool(field.get("italic", False))
        font_path  = field.get("fontPath")
        font_family = field.get("fontFamily")

        # Resolve local variant (if custom font path was provided)
        if font_path:
            font_path = _resolve_font_path(font_path, bold, italic, fonts_dir or output_dir)
        # Or download Google Font
        elif font_family and fonts_dir:
            font_path = _download_google_font(font_family, bold, italic, fonts_dir)

        try:
            fill_color = _hex_to_rgb(color_hex)
        except Exception:
            fill_color = (0, 0, 0)

        max_text_width = int(img_width * 0.85)
        font = _auto_shrink_font(draw, text, font_path, base_size, max_text_width)

        # Subtle shadow for readability if requested, but generally we draw exactly at (x, y)
        # Using anchor="mm" aligns the logical horizontal/vertical center exactly to the provided coords
        shadow = (0, 0, 0, 80) if img.mode == "RGBA" else (0, 0, 0)
        draw.multiline_text((x + 1, y + 1), text, font=font, fill=shadow, anchor="mm", align=alignment)
        draw.multiline_text((x, y), text, font=font, fill=fill_color, anchor="mm", align=alignment)

    # Flatten RGBA → RGB
    if img.mode == "RGBA":
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        img = bg

    name_value = _get_name_value(row, fields)
    safe_name  = _sanitize_filename(name_value) if name_value else "certificate"
    output_path = os.path.join(output_dir, f"{safe_name}.png")
    img.save(output_path, "PNG")
    return output_path


def _get_name_value(row: dict, fields: list[dict]) -> str:
    """Pick the 'name' field's value to use as the output filename."""
    for field in fields:
        col = field.get("column", "")
        if "name" in col.lower():
            val = str(row.get(col, "")).strip()
            if val and val.lower() != "nan":
                return val
    for field in fields:
        col = field.get("column", "")
        val = str(row.get(col, "")).strip()
        if val and val.lower() != "nan":
            return val
    return ""
