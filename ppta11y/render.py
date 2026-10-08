"""Pictures PowerPoint itself renders (Slide.Export), with shape ids drawn on, and before/after diffs."""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFont

from .com import safe
from .inventory import get_part, iter_shapes, shape_info, timeline_targets

WIDTH = 1600

COLORS = {
    "missing": (220, 30, 30),      # needs alt text and has none
    "junk": (170, 0, 170),         # alt text is a file name / URL / auto text
    "covered": (150, 150, 150),    # inside a group that has alt text or is decorative
    "review": (240, 140, 0),       # shape without text: alt text or decorative needed
    "ok": (20, 150, 60),           # has alt text
    "decorative": (130, 130, 130),
    "text": (40, 90, 200),         # read as text, no alt needed
}


def export(pres, slide: int, path: str | Path, width: int = WIDTH) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    h = round(width * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth)
    pres.Slides(slide).Export(str(path), "PNG", width, h)
    return path


def status(d: dict) -> str:
    from .ops import is_junk  # local import: ops imports render
    if d.get("covered"):
        return "covered"
    if d.get("decorative"):
        return "decorative"
    if d.get("alt", "").strip():
        return "junk" if is_junk(d["alt"]) else "ok"
    if d["needs_alt"] == "required":
        return "missing"
    if d["needs_alt"] == "review":
        return "review"
    return "text"


def _font(size: int):
    for name in ("segoeuib.ttf", "arialbd.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def labeled(pres, slide: int, out: str | Path, part: str = "slide", ids: list[int] | None = None,
            nested: bool = False, width: int = WIDTH) -> tuple[Path, list[dict]]:
    """Export the slide and draw a box + '#id' on each top-level shape (or on `ids`; `nested` also boxes
    shapes inside groups). For part='layout'/'master' the boxes are the inherited objects."""
    out = Path(out)
    export(pres, slide, out, width)
    scale = width / pres.PageSetup.SlideWidth
    container = get_part(pres, slide, part)
    targets, triggers = timeline_targets(container)
    chosen = []
    for s, depth, parent in iter_shapes(container.Shapes):
        if ids is not None:
            if s.Id not in ids:
                continue
        elif depth > 0 and not nested:
            continue
        if part != "slide" and s.Type == 14:
            continue
        d = shape_info(s, targets, triggers, deep=False) | {"depth": depth}
        if parent is not None and ((parent.AlternativeText or "").strip() or safe(lambda: parent.Decorative, 0)):
            d["covered"] = True
        d["rotation"] = safe(lambda: s.Rotation, 0) or 0
        chosen.append(d)

    img = Image.open(out).convert("RGB")
    over = Image.new("RGBA", img.size, (0, 0, 0, 0))
    dr = ImageDraw.Draw(over)
    font = _font(max(14, width // 80))
    for d in chosen:
        st = status(d)
        col = COLORS[st]
        x0, y0 = d["left"] * scale, d["top"] * scale
        x1, y1 = x0 + d["w"] * scale, y0 + d["h"] * scale
        dash = 2 if d["depth"] else 3
        if d["rotation"]:
            cx, cy, a = (x0 + x1) / 2, (y0 + y1) / 2, math.radians(d["rotation"])
            pts = [(cx + (x - cx) * math.cos(a) - (y - cy) * math.sin(a),
                    cy + (x - cx) * math.sin(a) + (y - cy) * math.cos(a))
                   for x, y in ((x0, y0), (x1, y0), (x1, y1), (x0, y1))]
            dr.polygon(pts, outline=col + (255,), width=dash)
        else:
            dr.rectangle([x0, y0, x1, y1], outline=col + (255,), width=dash)
        tag = f"#{d['id']}" + (" A" if d.get("animation_effects") else "")
        tb = dr.textbbox((0, 0), tag, font=font)
        tw, th = tb[2] - tb[0] + 8, tb[3] - tb[1] + 6
        lx, ly = max(0, min(x0, img.width - tw)), max(0, min(y0, img.height - th))
        dr.rectangle([lx, ly, lx + tw, ly + th], fill=col + (230,))
        dr.text((lx + 4, ly + 1), tag, fill=(255, 255, 255, 255), font=font)
        d["status"] = st
    img = Image.alpha_composite(img.convert("RGBA"), over).convert("RGB")
    img.save(out)
    return out, chosen


def numbered(pres, slide: int, out: str | Path, entries: list[tuple], width: int = WIDTH) -> Path:
    """Export the slide and draw a numbered badge on each (number, (l, t, r, b) in points, colour) entry:
    the order a screen reader reads the objects in."""
    out = Path(out)
    export(pres, slide, out, width)
    scale = width / pres.PageSetup.SlideWidth
    img = Image.open(out).convert("RGBA")
    over = Image.new("RGBA", img.size, (0, 0, 0, 0))
    dr = ImageDraw.Draw(over)
    font = _font(max(16, width // 70))
    for num, box, col in entries:
        x0, y0, x1, y1 = (v * scale for v in box)
        dr.rectangle([x0, y0, x1, y1], outline=col + (255,), width=3)
        tag = str(num)
        tb = dr.textbbox((0, 0), tag, font=font)
        tw, th = tb[2] - tb[0] + 14, tb[3] - tb[1] + 10
        lx, ly = max(0, min(x0, img.width - tw)), max(0, min(y0, img.height - th))
        dr.ellipse([lx, ly, lx + max(tw, th), ly + max(tw, th)], fill=col + (235,))
        dr.text((lx + (max(tw, th) - tw) / 2 + 7, ly + (max(tw, th) - th) / 2 + 3), tag, fill=(255, 255, 255, 255), font=font)
    Image.alpha_composite(img, over).convert("RGB").save(out)
    return out


def tolerant_changed(a: str | Path, b: str | Path, threshold: int = 24, radius: int = 1) -> int:
    """Pixels that changed even allowing a `radius`-pixel shift (antialiasing / sub-pixel line width):
    a pixel counts only if no pixel within `radius` in the other image has (nearly) its colour."""
    from PIL import ImageFilter
    A = Image.open(a).convert("RGB")
    B = Image.open(b).convert("RGB")
    if A.size != B.size:
        B = B.resize(A.size)
    k = 2 * radius + 1

    def unexplained(X, Y):  # pixels of Y outside the [min, max] of X's neighbourhood (per channel)
        lo, hi = X.filter(ImageFilter.MinFilter(k)), X.filter(ImageFilter.MaxFilter(k))
        out = None
        for ch in range(3):
            y, l_, h_ = Y.getchannel(ch), lo.getchannel(ch), hi.getchannel(ch)
            below = ImageChops.subtract(l_, y)  # l - y where y < l
            above = ImageChops.subtract(y, h_)
            m = ImageChops.lighter(below, above).point(lambda v: 255 if v > threshold else 0)
            out = m if out is None else ImageChops.lighter(out, m)
        return out
    m = ImageChops.lighter(unexplained(A, B), unexplained(B, A))
    return m.histogram()[255]


def diff(a: str | Path, b: str | Path, out: str | Path | None = None, threshold: int = 24) -> dict:
    """Fraction of pixels that changed between two exports, the changed area, and a side-by-side image."""
    A = Image.open(a).convert("RGB")
    B = Image.open(b).convert("RGB")
    if A.size != B.size:
        B = B.resize(A.size)
    r, g, b = ImageChops.difference(A, B).split()
    d = ImageChops.lighter(ImageChops.lighter(r, g), b).point(lambda v: 255 if v > threshold else 0)  # max channel
    changed = d.histogram()[255]
    res = {"changed_pixels": changed, "changed_fraction": round(changed / (A.width * A.height), 6),
           "changed_box_px": d.getbbox()}
    if out:
        W = A.width
        sheet = Image.new("RGB", (W * 2 + 20, A.height + 40), "white")
        sheet.paste(A, (0, 40))
        sheet.paste(B, (W + 20, 40))
        dr = ImageDraw.Draw(sheet)
        f = _font(24)
        dr.text((10, 6), "BEFORE", fill=(0, 0, 0), font=f)
        dr.text((W + 30, 6), "AFTER" + ("" if not res["changed_box_px"] else
                                         f"  (changed {res['changed_fraction']:.3%}, red box)"), fill=(0, 0, 0), font=f)
        if res["changed_box_px"]:
            x0, y0, x1, y1 = res["changed_box_px"]
            dr.rectangle([W + 20 + x0, 40 + y0, W + 20 + x1, 40 + y1], outline=(255, 0, 0), width=4)
        if sheet.width > 1800:  # keep inline images under the API's per-image size limit
            sheet = sheet.resize((1800, round(sheet.height * 1800 / sheet.width)))
        sheet.save(out)
        res["image"] = str(out)
    return res
