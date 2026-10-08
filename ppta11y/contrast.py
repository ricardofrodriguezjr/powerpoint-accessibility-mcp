"""Text contrast (WCAG 2.1 SC 1.4.3, AA): every run of text against the background it is drawn on.

  ratio      (L1 + 0.05) / (L2 + 0.05) with WCAG relative luminance; never rounded
  threshold  4.5:1, or 3:1 for large text (>= 18 pt, or >= 14 pt bold), applied to the rendered size: with
             "shrink text on overflow" PowerPoint's saved fontScale is used (half size assumed without it)
  background the fill of the table cell / shape holding the text if it is a solid, opaque colour;
             otherwise what PowerPoint actually renders behind the run (slide background, pictures,
             gradients, shapes underneath), sampled from Slide.Export
  fix        the nearest colour to the current text colour that passes (same hue, OKLCH lightness, chroma
             reduced only to stay in sRGB), keeping the text's light/dark polarity when possible. Text that
             is decorative, hidden, 100% transparent or entirely off the slide is exempt.
  states     text with an animated object over/under it is measured with and without that object; dimmed
             text and colour emphasis are measured in their animated colour
"""
from __future__ import annotations

import math
from collections import Counter

from PIL import Image

from . import render
from .com import ToolError, com_err, safe
from .inventory import get_part
from .session import checkpoint, out_dir, require_editable, revert, save, stamp

NORMAL, LARGE = 4.5, 3.0
MARGIN = 0.08          # aim this far above the threshold (8-bit rounding, checker differences)
SCALE = 2              # export at 2x for background sampling


# --------------------------------------------------------------------------------------------------
# Colour maths
# --------------------------------------------------------------------------------------------------

def bgr_to_rgb(v: int) -> tuple:
    return (v & 255, (v >> 8) & 255, (v >> 16) & 255)


def rgb_to_bgr(rgb) -> int:
    r, g, b = (int(round(x)) for x in rgb)
    return r + (g << 8) + (b << 16)


def hexrgb(rgb) -> str:
    return "#%02X%02X%02X" % tuple(int(round(x)) for x in rgb)


def _lin(c: float) -> float:
    c = c / 255
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def luminance(rgb) -> float:
    r, g, b = rgb
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def ratio(a, b) -> float:
    la, lb = luminance(a), luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def _to_oklab(rgb):
    r, g, b = (_lin(x) for x in rgb)
    l_ = (0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b) ** (1 / 3)
    m_ = (0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b) ** (1 / 3)
    s_ = (0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b) ** (1 / 3)
    return (0.2104542553 * l_ + 0.7936177850 * m_ - 0.0040720468 * s_,
            1.9779984951 * l_ - 2.4285922050 * m_ + 0.4505937099 * s_,
            0.0259040371 * l_ + 0.7827717662 * m_ - 0.8086757660 * s_)


def _from_oklab(L, a, b):
    l_ = L + 0.3963377774 * a + 0.2158037573 * b
    m_ = L - 0.1055613458 * a - 0.0638541728 * b
    s_ = L - 0.0894841775 * a - 1.2914855480 * b
    l, m, s = l_ ** 3, m_ ** 3, s_ ** 3
    lin = (4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s,
           -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s,
           -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s)

    def enc(c):
        return 12.92 * c if c <= 0.0031308 else 1.055 * (c ** (1 / 2.4)) - 0.055 if c > 0 else 12.92 * c
    return tuple(enc(c) * 255 for c in lin)


def _in_gamut(rgb, eps=0.5) -> bool:
    return all(-eps <= c <= 255 + eps for c in rgb)


def _lch_to_rgb(L, C, h):
    """OKLCH -> sRGB, reducing chroma (same lightness and hue) until it fits sRGB."""
    lo, hi = 0.0, C
    rgb = _from_oklab(L, C * math.cos(h), C * math.sin(h))
    if _in_gamut(rgb):
        return tuple(min(255, max(0, c)) for c in rgb)
    for _ in range(30):
        mid = (lo + hi) / 2
        rgb = _from_oklab(L, mid * math.cos(h), mid * math.sin(h))
        if _in_gamut(rgb):
            lo = mid
        else:
            hi = mid
    rgb = _from_oklab(L, lo * math.cos(h), lo * math.sin(h))
    return tuple(min(255, max(0, c)) for c in rgb)


def delta_e(a, b) -> float:
    """OKLab distance x100 (about 2 = just noticeable)."""
    A, B = _to_oklab(a), _to_oklab(b)
    return 100 * math.sqrt(sum((x - y) ** 2 for x, y in zip(A, B)))


def _min_ratio(c, bgs) -> float:
    return min(ratio(c, b) for b in bgs)


def nearest_passing(text_rgb, bg_rgb, target: float) -> tuple | None:
    """The closest colour to `text_rgb` (same hue, OKLCH lightness) whose contrast with the background is
    >= target (+ a small margin). `bg_rgb` is one colour or a list of colours the text sits on (all must
    pass). Keeps the light/dark polarity when possible. None if no single colour works."""
    bgs = [bg_rgb] if isinstance(bg_rgb[0], (int, float)) else list(bg_rgb)
    for want in (target + MARGIN, target + 0.005):  # margin first; the bare threshold if that is impossible
        if _min_ratio(text_rgb, bgs) >= want:
            return tuple(int(round(c)) for c in text_rgb)
        L0, a0, b0 = _to_oklab(text_rgb)
        C0, h = math.hypot(a0, b0), math.atan2(b0, a0)
        darker_first = luminance(text_rgb) <= min(luminance(b) for b in bgs) or \
            luminance(text_rgb) < sum(luminance(b) for b in bgs) / len(bgs)
        best = None
        for darker in ((True, False) if darker_first else (False, True)):
            extreme = (0, 0, 0) if darker else (255, 255, 255)
            if _min_ratio(extreme, bgs) < want:  # even pure black / white fails in this direction
                continue
            lo, hi = (0.0, L0) if darker else (L0, 1.0)
            ok, bad = (lo, hi) if darker else (hi, lo)  # `ok` is the end that passes
            for _ in range(40):
                mid = (ok + bad) / 2
                if _min_ratio(_lch_to_rgb(mid, C0, h), bgs) >= want:
                    ok = mid
                else:
                    bad = mid
            cand = tuple(int(round(c)) for c in _lch_to_rgb(ok, C0, h))
            step = 0
            while _min_ratio(cand, bgs) < want and step < 256:  # 8-bit rounding: one code value at a time
                cand = tuple(max(0, c - 1) for c in cand) if darker else tuple(min(255, c + 1) for c in cand)
                step += 1
            if _min_ratio(cand, bgs) < want:
                cand = extreme
            if best is None or delta_e(cand, text_rgb) < delta_e(best, text_rgb):
                best = cand
            if darker == darker_first:
                break  # same polarity found: prefer it
        if best is not None:
            return best
    return None


# --------------------------------------------------------------------------------------------------
# Finding text and its background
# --------------------------------------------------------------------------------------------------

def _brighten(rgb, b):
    import colorsys
    h, l, s = colorsys.rgb_to_hls(*(c / 255 for c in rgb))
    l = l * (1 + b) if b < 0 else l + (1 - l) * b
    return tuple(c * 255 for c in colorsys.hls_to_rgb(h, l, s))


def _num(v, default=0.0) -> float:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return default
    return v if math.isfinite(v) and abs(v) < 100000 else default


def _effective(color_format, rgb):
    """ColorFormat.RGB is the base colour; theme tints/shades (Brightness, TintAndShade) are applied here."""
    br = _num(safe(lambda: color_format.Brightness, 0))
    if br and safe(lambda: color_format.ObjectThemeColor, 0):
        rgb = _brighten(rgb, br)
    ts = _num(safe(lambda: color_format.TintAndShade, 0))
    if ts:
        rgb = _brighten(rgb, ts)
    return rgb


def _solid_fill(s):
    f = safe(lambda: s.Fill)
    if f is None or safe(lambda: f.Visible, 0) != -1 or safe(lambda: f.Type) != 1:
        return None
    if (safe(lambda: f.Transparency, 0) or 0) > 0.001:
        return None
    v = safe(lambda: f.ForeColor.RGB)
    if not isinstance(v, int) or v < 0:
        return None
    return tuple(round(c) for c in _effective(f.ForeColor, bgr_to_rgb(v)))


def _text_holders(sl):
    """(part, top-level shape, text-holding shape, label, cell-or-node or None, frame id) for every piece of
    visible text: slide shapes, group children, table cells (merged cells once), SmartArt nodes, and the
    non-placeholder text of the layout/master the slide shows. Charts come with holder None (not measured).
    `frame id` is the id of the shape holding the text, table or SmartArt (which can sit inside a group)."""
    def walk(s, top, part, prefix):
        if safe(lambda: s.Decorative, 0) == -1 or not safe(lambda: s.Visible, -1):
            return
        if s.Type == 6:
            for i in range(1, s.GroupItems.Count + 1):
                yield from walk(s.GroupItems(i), top, part, prefix)
            return
        inside = f" in group #{top.Id}" if top.Id != s.Id else ""
        if safe(lambda: s.HasChart) == -1:
            yield part, top, None, f"{prefix}#{s.Id} chart{inside}", None, s.Id
            return
        if safe(lambda: s.HasSmartArt) == -1:
            nodes = safe(lambda: s.SmartArt.AllNodes)
            for k in range(1, (safe(lambda: nodes.Count, 0) or 0) + 1):
                shp = safe(lambda k=k: nodes.Item(k).Shapes)
                for j in range(1, (safe(lambda: shp.Count, 0) or 0) + 1):
                    ns = shp.Item(j)
                    if safe(lambda: ns.HasTextFrame) == -1 and safe(lambda: ns.TextFrame.HasText):
                        yield part, top, ns, f"{prefix}#{s.Id} SmartArt node {k}{inside}", ("smartart", k, j), s.Id
            return
        if safe(lambda: s.HasTable) == -1:
            t = s.Table
            seen = set()
            for r in range(1, t.Rows.Count + 1):
                for c in range(1, t.Columns.Count + 1):
                    cs = t.Cell(r, c).Shape
                    key = (round(cs.Left, 1), round(cs.Top, 1), round(cs.Width, 1), round(cs.Height, 1))
                    if key in seen:  # merged cells report the same shape at every position
                        continue
                    seen.add(key)
                    yield part, top, cs, f"{prefix}#{s.Id} cell R{r}C{c}{inside}", (r, c), s.Id
            return
        if safe(lambda: s.HasTextFrame) == -1 and safe(lambda: s.TextFrame.HasText):
            yield part, top, s, f"{prefix}#{s.Id}{inside}", None, s.Id
    for i in range(1, sl.Shapes.Count + 1):
        yield from walk(sl.Shapes(i), sl.Shapes(i), "slide", "")
    if safe(lambda: sl.DisplayMasterShapes, -1):
        lay = sl.CustomLayout
        parts = [("layout", lay, "layout ")]
        if safe(lambda: lay.DisplayMasterShapes, -1):
            parts.append(("master", sl.Master, "master "))
        for part, cont, prefix in parts:
            for i in range(1, cont.Shapes.Count + 1):
                s = cont.Shapes(i)
                if s.Type != 14:  # placeholders on layouts/masters are prompts, never shown on slides
                    yield from walk(s, s, part, prefix)


def _text_fill(font2):
    """(colour before transparency, opacity) of a run's text; "pattern" for gradient / pattern / picture
    text; None when the text is not drawn (no fill, or fully transparent)."""
    fill = font2.Fill
    if safe(lambda: fill.Visible, -1) == 0:
        return None
    a = 1 - _num(safe(lambda: fill.Transparency, 0))
    if a <= 0.01:
        return None  # 100% transparent text is not shown (e.g. a screen-reader-only label): exempt
    if safe(lambda: fill.Type, 1) in (2, 3, 4, 6):
        return "pattern"
    v = safe(lambda: fill.ForeColor.RGB)
    if not isinstance(v, int) or v < 0:
        return None
    return _effective(fill.ForeColor, bgr_to_rgb(v)), a


def _blend(tf, bg) -> tuple:
    rgb, a = tf
    if a < 1 and bg is not None:
        return tuple(a * c + (1 - a) * b for c, b in zip(rgb, bg))
    return tuple(rgb)


def _sample_bg(img, box_pt, scale, text_rgb, share: float = 0.08, need_glyphs: bool = False) -> list:
    """The background colours actually rendered right next to the letters (WCAG G18: the area immediately
    around each glyph, sampled just outside the antialiased edge): every cluster of such pixels holding at
    least `share` of them (mean colour of the cluster's real pixels). Text over a gradient, picture or split background gets several; the run must pass
    against all of them. Falls back to the whole run box when the glyphs cannot be found."""
    from PIL import ImageFilter
    x0, y0, x1, y1 = (int(v * scale) for v in box_pt)
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(img.width, max(x0 + 1, x1)), min(img.height, max(y0 + 1, y1))
    if x1 <= x0 or y1 <= y0:
        return []
    crop = img.crop((x0, y0, x1, y1))
    px = list(crop.getdata())
    tr = text_rgb or (0, 0, 0)
    is_text = [sum(abs(a - b) for a, b in zip(p, tr)) <= 90 for p in px]
    pool, fallback = [], False
    if sum(is_text) >= 4:
        mask = Image.new("L", crop.size)
        mask.putdata([255 if t else 0 for t in is_text])
        # a ring 3-4 px (at 2x) around the letters: close enough to be what the letters are seen against,
        # outside the 1-2 px band of antialiased edge pixels (blends of text and background colour)
        outer = list(mask.filter(ImageFilter.MaxFilter(9)).getdata())
        inner = list(mask.filter(ImageFilter.MaxFilter(5)).getdata())
        pool = [p for p, o, i in zip(px, outer, inner) if o and not i]
    if len(pool) < 20:
        if need_glyphs:  # the letters cannot be seen here (e.g. under a cover): nothing to measure
            return []
        far = [p for p, t in zip(px, is_text) if not t]
        if len(far) >= 0.2 * len(px):
            pool = far
        else:  # text and background too alike to tell apart: every pixel, glyphs included
            pool, fallback = px, True
    buckets: dict = {}
    for p in pool:
        buckets.setdefault((p[0] // 12, p[1] // 12, p[2] // 12), []).append(p)
    if not buckets:
        return []
    ranked = sorted(buckets.values(), key=len, reverse=True)
    keep = [b for b in ranked if len(b) >= share * len(pool)] or ranked[:1]
    means = [tuple(sum(c[i] for c in b) / len(b) for i in range(3)) for b in keep]
    if fallback:  # the glyphs themselves form a cluster: it is the text, not its background
        other = [m for m in means if max(abs(a - b) for a, b in zip(m, tr)) > 8]
        if other:
            means = other
    return means


def _runs(holder):
    tr2 = holder.TextFrame2.TextRange
    return [tr2.Runs.Item(i) for i in range(1, tr2.Runs.Count + 1)]


_DECK_XML: dict = {}


def _timing_model(root, P) -> tuple[list[dict], set]:
    """Which shapes are on screen at each step of the slide's main animation sequence, from the saved XML:
    ([visibility after step 0, 1, 2, ... as {shape id: shown}], ids an interactive (trigger) sequence shows
    or hides at any time). Shapes no entrance/exit effect targets are always shown."""
    main, inter = [], []
    for seq in root.iter(f"{P}seq"):
        ctn = seq.find(f"{P}cTn")
        lst = ctn.find(f"{P}childTnLst") if ctn is not None else None
        if lst is None:
            continue
        nt = ctn.get("nodeType")
        groups = [list(par.iter(f"{P}cTn")) for par in lst.findall(f"{P}par")] if nt == "mainSeq" else \
            [list(ctn.iter(f"{P}cTn"))]
        for cts in groups:
            step = []
            for c in cts:
                cls = c.get("presetClass")
                if cls not in ("entr", "exit"):
                    continue
                tgt = c.find(f".//{P}spTgt")
                if tgt is not None and tgt.get("spid", "").isdigit():
                    step.append((int(tgt.get("spid")), cls))
            (main.append(step) if nt == "mainSeq" else inter.extend(step))
    first: dict = {}
    for sid, cls in [x for st in main for x in st] + inter:
        first.setdefault(sid, cls != "entr")  # an entrance target starts hidden; an exit target starts shown
    cur, states = dict(first), [dict(first)]
    for st in main:
        for sid, cls in st:
            cur[sid] = cls == "entr"
        states.append(dict(cur))
    return states, {sid for sid, _ in inter}


def deck_xml(pres, refresh: bool = False) -> dict:
    """Facts only the saved file has, read from one saved copy and cached per deck until refresh:
    "scales" {(SlideID, shape id): fontScale} for text PowerPoint has shrunk ("shrink text on overflow" writes
    <a:normAutofit fontScale="62500"/>), and "timing" {SlideID: _timing_model}."""
    import os
    import zipfile
    import xml.etree.ElementTree as ET
    from .animations import slide_part_name
    from .session import save_copy, tmp_dir
    key = pres.FullName
    if key in _DECK_XML and not refresh:
        return _DECK_XML[key]
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    out = {"scales": {}, "timing": {}}
    tmp = str(tmp_dir(pres) / f"fs{os.getpid() % 10000}.pptx")
    try:
        save_copy(pres, tmp)
        with zipfile.ZipFile(tmp) as z:
            for n in range(1, pres.Slides.Count + 1):
                sid = pres.Slides(n).SlideID
                root = ET.fromstring(z.read(slide_part_name(z, sid)))
                for sp in root.iter(f"{P}sp"):
                    na = sp.find(f"{P}txBody/{A}bodyPr/{A}normAutofit")
                    cnv = sp.find(f"{P}nvSpPr/{P}cNvPr")
                    if na is not None and cnv is not None and na.get("fontScale"):
                        out["scales"][(sid, int(cnv.get("id")))] = int(na.get("fontScale")) / 100000
                out["timing"][sid] = _timing_model(root, P)
    except Exception:  # noqa: BLE001 - no file facts: conservative rules apply (half size; every state)
        out = {"scales": None, "timing": None}
    finally:
        safe(lambda: os.remove(tmp))
    _DECK_XML[key] = out
    return out


def font_scales(pres, refresh: bool = False) -> dict | None:
    return deck_xml(pres, refresh)["scales"]


def _patterns(model, t_ids: set, near: list[int]) -> list[frozenset]:
    """The sets of nearby animated objects that are hidden while the text is on screen, one per distinct
    situation of the show (each step of the main sequence; trigger-driven objects either way). Without a
    model: as shown, and with every nearby animated object hidden."""
    if model is None:
        return [frozenset(), frozenset(near)] if near else [frozenset()]
    states, toggles = model
    pats = []
    for st in states:
        if not all(st.get(x, True) for x in t_ids) and not (t_ids & toggles):
            continue  # the text itself is not on screen at this step
        hidden = frozenset(o for o in near if not st.get(o, True))
        cands = [hidden] + [hidden ^ {o} for o in near if o in toggles]
        pats += [c for c in cands if c not in pats]
    return pats or [frozenset()]


# Emphasis effects that change how text looks during the show. 56/54 are measured (their target colour is
# known); the others blend or flash colours in ways that cannot be computed: listed for a check by eye.
FONT_COLOR_EFFECT, FILL_COLOR_EFFECT = 56, 54
COLOR_EFFECTS = {11, 27, 62, 63, 66, 68, 69, 70, 71, 72, 73, 74, 75, 77, 78}


def _anim_info(sl, slide_box):
    """(animated objects as [(top index, object id, painted box, shape)], other text states per shape id as
    [(state, kind, rgb or None, paragraph or 0)]), from the slide's main and interactive sequences."""
    from .inventory import timeline_targets
    from .readingorder import _moving_ids, painted_box
    targets, triggers = timeline_targets(sl)
    anim = set(targets) | set(triggers)
    objs, states = [], {}
    if not anim:
        return objs, states
    moving = _moving_ids(sl)
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        if not safe(lambda: s.Visible, -1):
            continue
        cands = [s] + ([s.GroupItems(k) for k in range(1, s.GroupItems.Count + 1)] if s.Type == 6 else [])
        for o in cands:
            if o.Id in anim:
                objs.append((i, o.Id, painted_box(o, slide_box, moving=o.Id in moving), o))
    tl = sl.TimeLine
    seqs = [tl.MainSequence] + [tl.InteractiveSequences(k) for k in range(1, tl.InteractiveSequences.Count + 1)]
    for q in seqs:
        for k in range(1, q.Count + 1):
            e = q.Item(k)
            sid, et = safe(lambda: e.Shape.Id), safe(lambda: e.EffectType)
            para = safe(lambda: e.Paragraph, 0) or 0
            info = safe(lambda: e.EffectInformation)
            if safe(lambda: info.AfterEffect) == 1:
                v = safe(lambda: info.Dim.RGB)
                states.setdefault(sid, []).append(("dimmed after its animation", "text",
                                                   bgr_to_rgb(v) if isinstance(v, int) and v >= 0 else None, para))
            if et in (FONT_COLOR_EFFECT, FILL_COLOR_EFFECT):
                to = None
                for b in range(1, (safe(lambda: e.Behaviors.Count, 0) or 0) + 1):
                    beh = e.Behaviors.Item(b)
                    if safe(lambda: beh.Type) == 2:  # msoAnimTypeColor
                        v = safe(lambda: beh.ColorEffect.To.RGB)
                        to = bgr_to_rgb(v) if isinstance(v, int) and v >= 0 else None
                states.setdefault(sid, []).append(
                    ("font colour emphasis" if et == FONT_COLOR_EFFECT else "fill colour emphasis",
                     "text" if et == FONT_COLOR_EFFECT else "fill", to, para))
            elif et in COLOR_EFFECTS:
                states.setdefault(sid, []).append((f"colour emphasis effect {et}", "other", None, para))
    return objs, states


def _render_without(pres, slide: int, objs: list, path) -> Image.Image:
    """The slide rendered with the given animated objects hidden (the state before an entrance / after an
    exit); they are shown again afterwards and the deck's saved state is kept."""
    was_saved = bool(pres.Saved)
    hidden = []
    try:
        for o in objs:
            if safe(lambda: o.Visible, -1):
                o.Visible = 0
                hidden.append(o)
        render.export(pres, slide, path, int(pres.PageSetup.SlideWidth * SCALE))
    finally:
        for o in hidden:
            o.Visible = -1
        if was_saved:
            pres.Saved = -1
    return Image.open(path).convert("RGB")


def _paragraph_starts(holder) -> list[int]:
    tr2 = holder.TextFrame2.TextRange
    return [tr2.Paragraphs.Item(i).Start for i in range(1, tr2.Paragraphs.Count + 1)]


def slide_contrast(pres, slide: int, img=None, scales: dict | None = None, only: set | None = None) -> list[dict]:
    """Every run of visible text on the slide (and the layout/master text it shows) with its contrast
    against the worst background colour it is drawn on. Text near an animated object is measured in every
    situation of the show in which the text is on screen (the object shown or not, step by step, from the
    saved animation sequence). Other states of the text (dimmed after its animation, colour emphasis) are
    extra rows with "state". Chart text and colour effects that cannot be computed: rows with passes=None.
    `only` = {(part, shape id)} limits the holders measured."""
    from .readingorder import _rotate_rect, overlaps
    sl = get_part(pres, slide, "slide")
    facts = deck_xml(pres)
    if scales is None:
        scales = facts["scales"]
    model = (facts["timing"] or {}).get(sl.SlideID) if facts["timing"] is not None else None
    W, H = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
    if img is None:
        p = out_dir(pres) / "tmp" / f"c{slide}.png"
        render.export(pres, slide, p, int(W * SCALE))
        img = Image.open(p).convert("RGB")
    scale = img.width / W
    objs, states = _anim_info(sl, (0.0, 0.0, W, H))
    zpos = {sl.Shapes(i).Id: i for i in range(1, sl.Shapes.Count + 1)}
    out, pending = [], []
    for part, top, holder, label, cell, frame in _text_holders(sl):
        if only is not None and (part, top.Id) not in only:
            continue
        base = {"slide": slide, "part": part, "where": label, "shape_id": top.Id, "frame_id": frame,
                "holder_id": safe(lambda: holder.Id) if holder is not None and cell is None else None, "cell": cell}
        if holder is None:
            out.append({**base, "text": "(chart text)", "passes": None,
                        "note": "chart text is not measured; check the chart's colours by eye"})
            continue
        fill_bg = _solid_fill(holder)
        shrink = safe(lambda: holder.TextFrame2.AutoSize, 0) == 2
        hid = base["holder_id"]
        my_states = [] if part != "slide" else states.get(top.Id, []) + (states.get(hid, []) if hid != top.Id else [])
        pstarts = _paragraph_starts(holder) if my_states else []
        for r in _runs(holder):
            txt = (r.Text or "").replace("\r", " ").replace("\x0b", " ").strip()
            if not txt:
                continue
            bl, bt, bw, bh = (safe(lambda: r.BoundLeft), safe(lambda: r.BoundTop), safe(lambda: r.BoundWidth),
                              safe(lambda: r.BoundHeight))
            if None in (bl, bt, bw, bh) or bw <= 0 or bh <= 0:
                continue
            if cell:  # text bounds of table cells / SmartArt nodes are relative to the cell: sample the cell
                area = (holder.Left, holder.Top, holder.Left + holder.Width, holder.Top + holder.Height)
            else:  # bounds are for the unrotated frame
                area = _rotate_rect((bl, bt, bl + bw, bt + bh), holder)
            if area[2] <= 0 or area[3] <= 0 or area[0] >= W or area[1] >= H:
                continue  # off-slide (e.g. a hidden title): never shown
            font = r.Font
            tf = _text_fill(font)
            if tf is None:
                continue
            row = {**base, "run_start": r.Start, "run_length": r.Length, "text": txt[:50]}
            if tf == "pattern":
                out.append({**row, "passes": None, "note": "gradient / pattern / picture text fill is not measured; "
                                                            "check it by eye (or give the text a solid colour)"})
                continue
            covers = [o for o in objs if o[1] not in (top.Id, hid) and overlaps(o[2], area)]
            above = {o[1] for o in covers if part != "slide" or o[0] > zpos.get(top.Id, 0)}
            t_ids = {x for x in (top.Id, hid, frame) if x is not None} if part == "slide" else set()
            pats = _patterns(model, t_ids, [o[1] for o in covers]) if covers else [frozenset()]
            size = safe(lambda: font.Size, 0) or 0
            bold = safe(lambda: font.Bold, 0) == -1
            # "shrink text on overflow" renders text smaller than its nominal size: use PowerPoint's own
            # saved scale; without it, assume it could be half size
            eff = size
            if shrink:
                sc = scales.get((sl.SlideID, hid)) if scales is not None and part == "slide" and cell is None else None
                eff = size * (sc if sc is not None else (1.0 if scales is not None and part == "slide" and cell is None
                                                          else 0.5))
            para = next((i for i, s0 in reversed(list(enumerate(pstarts, 1))) if s0 <= r.Start), 1) if pstarts else 0
            pending.append({"row": row, "tf": tf, "area": area, "fill_bg": fill_bg, "covers": covers, "above": above,
                            "pats": pats, "size": size, "eff": eff, "bold": bold, "para": para, "states": my_states})
    shapes = {o[1]: o[3] for o in objs}
    images = {frozenset(): img}
    for pat in {pt for pd in pending for pt in pd["pats"]} - {frozenset()}:
        images[pat] = _render_without(pres, slide, [shapes[i] for i in sorted(pat)],
                                      out_dir(pres) / "tmp" / f"c{slide}_state{len(images)}.png")
    for pd in pending:
        raw = pd["tf"][0]
        samples, refs = [], []
        for pat in pd["pats"]:
            # an animated object shown over the text at this step can hide it: count only text actually seen
            covered = bool(pd["above"] - pat)
            got = _sample_bg(images[pat], pd["area"], scale, raw, need_glyphs=covered)
            samples += [g for g in got if g not in samples]
            refs = refs or got
        fill_bg = pd["fill_bg"]
        if fill_bg is not None and (not refs or max(abs(a - b) for a, b in zip(fill_bg, refs[0])) <= 12):
            bgs, src = [fill_bg], "fill"
        else:  # no solid fill, or the fill does not render as its model colour: trust the pixels
            bgs, src = samples, "render"
        if not bgs:
            continue
        large = pd["eff"] >= 18.0 or (pd["bold"] and pd["eff"] >= 14.0)
        need = LARGE if large else NORMAL
        common = {"size": pd["size"], **({"rendered_size": round(pd["eff"], 1)} if pd["eff"] != pd["size"] else {}),
                  "bold": pd["bold"], "needs": need, "bg_from": src,
                  **({"measured_states": [("as shown" if not pt else "without #" + ", #".join(map(str, sorted(pt))))
                                          for pt in pd["pats"]],
                      "animated_neighbours": [o[1] for o in pd["covers"]]} if pd["covers"] else {})}
        colors = [_blend(pd["tf"], b) for b in bgs]
        rts = [ratio(c, b) for c, b in zip(colors, bgs)]
        k = rts.index(min(rts))
        out.append({**pd["row"], **common, "text_rgb": hexrgb(colors[k]), "bg_rgb": hexrgb(bgs[k]),
                    "bg_all": [hexrgb(b) for b in bgs], "ratio": round(rts[k], 2), "passes": rts[k] >= need})
        for state, what, rgb, para in pd["states"]:
            if para and pd["para"] and para != pd["para"]:
                continue
            srow = {**pd["row"], **common, "state": state}
            if rgb is None or what == "other":
                out.append({**srow, "passes": None, "note": f"{state}: the text's colour changes during the show in "
                                                           "a way that is not measured; check it by eye"})
                continue
            if what == "fill":
                if fill_bg is None:
                    continue  # the effect recolours a fill this text does not sit on
                sb = [rgb]
                sc = [_blend(pd["tf"], rgb)]
            else:
                sb = bgs
                sc = [_blend((rgb, pd["tf"][1]), b) for b in bgs]
            sr = [ratio(c, b) for c, b in zip(sc, sb)]
            j = sr.index(min(sr))
            out.append({**srow, "text_rgb": hexrgb(sc[j]), "bg_rgb": hexrgb(sb[j]), "bg_all": [hexrgb(b) for b in sb],
                        "ratio": round(sr[j], 2), "passes": sr[j] >= need})
    return out


def contrast_report(pres, slides: list[int] | None = None) -> dict:
    slides = slides or list(range(1, pres.Slides.Count + 1))
    rows = []
    for n in slides:
        rows += slide_contrast(pres, n)
    fails = [r for r in rows if r["passes"] is False]
    unchecked = [r for r in rows if r["passes"] is None]
    return {"runs_checked": len(rows) - len(unchecked), "failing": len(fails), "failures": fails,
            "slides_with_failures": sorted({r["slide"] for r in fails}),
            "not_machine_checked": [{"slide": r["slide"], "where": r["where"], "note": r["note"],
                                     **({"text": r["text"]} if r.get("text") else {})} for r in unchecked]}


# --------------------------------------------------------------------------------------------------
# Fix
# --------------------------------------------------------------------------------------------------

def _container(sl, part):
    return sl if part == "slide" else sl.CustomLayout if part == "layout" else sl.Master


def _find_in(s, sid: int):
    if s.Id == sid:
        return s
    if s.Type == 6:
        for i in range(1, s.GroupItems.Count + 1):
            f = _find_in(s.GroupItems(i), sid)
            if f is not None:
                return f
    return None


def _holder(sl, row):
    cont = _container(sl, row.get("part", "slide"))
    s = None
    for i in range(1, cont.Shapes.Count + 1):
        if cont.Shapes(i).Id == row["shape_id"]:
            s = cont.Shapes(i)
    if s is None:
        raise ToolError(f"shape #{row['shape_id']} is gone")
    cell = row["cell"]
    if cell:
        f = _find_in(s, row.get("frame_id") or s.Id)  # a table / SmartArt can sit inside a group
        if f is None:
            raise ToolError(f"table/SmartArt #{row.get('frame_id')} not found in #{s.Id}")
        if cell[0] == "smartart":
            return f.SmartArt.AllNodes.Item(cell[1]).Shapes.Item(cell[2])
        return f.Table.Cell(*cell).Shape
    h = _find_in(s, row["holder_id"])
    if h is None:
        raise ToolError(f"text holder #{row['holder_id']} not found")
    return h


def _rgb(h: str) -> tuple:
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def _same_fill_holders(sl, rgb: tuple) -> list:
    """Every slide shape, group child and table cell (with text or not) whose solid fill is `rgb`."""
    out = []

    def walk(s):
        if s.Type == 6:
            for i in range(1, s.GroupItems.Count + 1):
                walk(s.GroupItems(i))
            return
        if safe(lambda: s.HasTable) == -1:
            t = s.Table
            for r in range(1, t.Rows.Count + 1):
                for c in range(1, t.Columns.Count + 1):
                    cs = t.Cell(r, c).Shape
                    if _solid_fill(cs) == rgb:
                        out.append(cs)
            return
        if _solid_fill(s) == rgb:
            out.append(s)
    for i in range(1, sl.Shapes.Count + 1):
        walk(sl.Shapes(i))
    return out


def _fill_eligible(r: dict) -> bool:
    """Rows whose background fill _same_fill_holders can recolour: slide shapes and table cells (not
    SmartArt nodes, not layout/master objects), measured on their fill, in their normal state."""
    cell = r.get("cell")
    return (r.get("bg_from") == "fill" and r.get("part", "slide") == "slide" and not r.get("state")
            and not (cell and cell[0] == "smartart"))


def _fill_plans(rows_all: list[dict], fails: list[dict]) -> dict:
    """{fill hex: new fill} where darkening/lightening a shared slide cell/shape fill (same hue) is a much
    smaller visible change than recolouring the text (e.g. white text on a mid-tone fill would otherwise
    turn black). Every run on that fill, failing or not, must pass with the new fill."""
    plans = {}
    for fill in {r["bg_rgb"] for r in fails if _fill_eligible(r)}:
        on = [r for r in rows_all if _fill_eligible(r) and r.get("bg_rgb") == fill]
        bad = [r for r in fails if _fill_eligible(r) and r["bg_rgb"] == fill]
        new = nearest_passing(_rgb(fill), [_rgb(r["text_rgb"]) for r in bad], max(r["needs"] for r in bad))
        if new is None or any(ratio(_rgb(r["text_rgb"]), new) < r["needs"] + MARGIN / 2 for r in on):
            continue
        de_fill = delta_e(_rgb(fill), new)
        de_text = max(delta_e(_rgb(r["text_rgb"]), nearest_passing(_rgb(r["text_rgb"]), _rgb(fill), r["needs"])
                               or (0, 0, 0)) for r in bad)
        if de_fill * 2 < de_text:
            plans[fill] = new
    return plans


def _shown_on(pres, sl, part: str) -> list[int]:
    from .inventory import slides_showing_master, slides_using_layout
    if part == "layout":
        return slides_using_layout(pres, sl.CustomLayout, shown_only=True)
    return slides_showing_master(pres, sl.Design.Index)


def _same_run(a: dict, b: dict) -> bool:
    return all(a.get(k) == b.get(k) for k in ("part", "shape_id", "holder_id", "cell", "frame_id", "run_start")) \
        and not a.get("state") and a.get("passes") is not None


def _state_advice(r: dict) -> str:
    bgs = [_rgb(b) for b in r.get("bg_all") or [r["bg_rgb"]]]
    new = nearest_passing(_rgb(r["text_rgb"]), bgs, r["needs"])
    how = {"dimmed after its animation": "Animation Pane > Effect Options > After animation: pick a dim colour",
           "font colour emphasis": "Animation Pane > Effect Options: pick a font colour",
           "fill colour emphasis": "Animation Pane > Effect Options: pick a fill colour"}.get(r["state"], "")
    return (f"{r['state']}: the colour is stored in the animation, which this tool does not recolour. {how}"
            + (f" such as {hexrgb(new)} (passes {_min_ratio(new, bgs):.2f}:1)" if new and r["state"] != "fill colour emphasis" else ""))


def _fix_slide(pres, n: int, rows_all: list[dict], strategy: str, out: dict) -> list[dict]:
    """Apply the fixes for one slide (no save). Returns the changes; raises on failure. Text on the slide's
    layout/master is shared with other slides: its new colour must pass on every slide that shows it, and
    is re-measured there."""
    rows = [r for r in rows_all if r["passes"] is False]
    for r in [r for r in rows if r.get("state")]:
        out["unfixable"].append({**r, "why": _state_advice(r)})
    rows = [r for r in rows if not r.get("state")]
    plans = _fill_plans(rows_all, rows) if strategy == "auto" else {}
    sl = get_part(pres, n, "slide")
    holders = {fill: _same_fill_holders(sl, _rgb(fill)) for fill in plans}
    plans = {f: v for f, v in plans.items() if holders[f]}
    changed, verify = [], {}
    for fill, new in plans.items():
        for h in holders[fill]:
            h.Fill.ForeColor.RGB = rgb_to_bgr(new)
            safe(lambda: setattr(h.Fill.ForeColor, "TintAndShade", 0))
        ons = [x for x in rows_all if x.get("bg_rgb") == fill and _fill_eligible(x)]
        changed.append({"slide": n, "where": ", ".join(sorted({x["where"] for x in ons}))[:200],
                        "text": "; ".join(sorted({x["text"] for x in ons}))[:120], "fill_from": fill,
                        "fill_to": hexrgb(new), "text_color": ", ".join(sorted({x["text_rgb"] for x in ons})),
                        "shapes_recoloured": len(holders[fill]),
                        "ratio_before": min(x["ratio"] for x in ons),
                        "ratio_after": round(min(ratio(_rgb(x["text_rgb"]), new) for x in ons), 2),
                        "delta_e": round(delta_e(_rgb(fill), new), 1), "changed": "fill"})
    for r in [r for r in rows if not (_fill_eligible(r) and r["bg_rgb"] in plans)]:
        old = _rgb(r["text_rgb"])
        bgs = {b for b in r.get("bg_all") or [r["bg_rgb"]]}
        part = r.get("part", "slide")
        shown = [n]
        if part != "slide":  # shared object: every slide that shows it decides the colour
            shown = _shown_on(pres, sl, part) or [n]
            for m in shown:
                if m != n:
                    bgs |= {b for x in slide_contrast(pres, m, only={(part, r["shape_id"])}) if _same_run(x, r)
                            for b in x.get("bg_all") or [x["bg_rgb"]]}
            verify.setdefault((part, r["shape_id"]), set()).update(m for m in shown if m != n)
        bgs = [_rgb(b) for b in sorted(bgs)]
        new = nearest_passing(old, bgs, r["needs"])
        if new is None:
            out["unfixable"].append({**r, "why": "no single text colour reaches the ratio on this background "
                                                 "(text over a mixed light/dark area" +
                                                 (f", or a {part} object shown on differently coloured slides"
                                                  if part != "slide" else "") + "): needs a backing shape"})
            continue
        h = _holder(sl, r)
        rng = h.TextFrame2.TextRange.GetCharacters(r["run_start"], r["run_length"])
        rng.Font.Fill.ForeColor.RGB = rgb_to_bgr(new)
        safe(lambda: setattr(rng.Font.Fill.ForeColor, "TintAndShade", 0))
        safe(lambda: setattr(rng.Font.Fill.ForeColor, "Brightness", 0))
        if (safe(lambda: rng.Font.Fill.Transparency, 0) or 0) > 0:
            rng.Font.Fill.Transparency = 0
        changed.append({"slide": n, "where": r["where"], "part": part, "text": r["text"],
                        "from": r["text_rgb"], "to": hexrgb(new), "on": ", ".join(hexrgb(b) for b in bgs),
                        "ratio_before": r["ratio"], "ratio_after": round(_min_ratio(new, bgs), 2),
                        "needs": r["needs"], "delta_e": round(delta_e(old, new), 1), "changed": "text colour",
                        **({"shown_on_slides": shown} if part != "slide" else {})})
    for (part, sid), others in verify.items():  # the shared object, re-measured where else it shows
        for m in sorted(others):
            bad = [x for x in slide_contrast(pres, m, only={(part, sid)}) if x["passes"] is False and not x.get("state")]
            if bad:
                raise ToolError(f"{part} text #{sid} still fails on slide {m} after the fix: "
                                f"{[(x['text'], x['ratio']) for x in bad][:3]}")
    return changed


def fix_contrast(pres, slides: list[int] | None = None, strategy: str = "auto", autosave: bool = True) -> dict:
    """Make every failing run pass with the smallest visible change: the nearest passing text colour (same
    hue), or - strategy "auto" - a slightly darker/lighter cell or shape fill when that changes much less
    (every shape/cell with that fill is recoloured; all text on it re-checked). strategy="text" never
    touches fills. One transaction per slide (a failing "auto" slide is retried with "text"); every change
    is re-measured (layout/master text on every slide that shows it). Colours PowerPoint keeps inside an
    animation (dim after animation, colour emphasis) cannot be changed by a tool: they are returned in
    `unfixable` with the colour to pick."""
    from . import session
    require_editable(pres)
    path = pres.FullName
    slides = slides or list(range(1, pres.Slides.Count + 1))
    out = {"changed": [], "failed": {}, "unfixable": [], "not_machine_checked": []}
    for n in slides:
        attempts = [strategy] if strategy == "text" else [strategy, "text"]
        for strat in attempts:
            p = session.find_open(path)
            try:
                rows_all = slide_contrast(p, n)
            except Exception as exc:  # noqa: BLE001
                out["failed"][n] = [f"measuring: {type(exc).__name__}: {com_err(exc)}"]
                break
            if strat == attempts[0]:
                out["not_machine_checked"] += [r for r in rows_all if r["passes"] is None]
            if not any(r["passes"] is False for r in rows_all):
                break
            checkpoint(p)
            failures, changed = [], []
            unfix_before = len(out["unfixable"])
            try:
                od, tag = out_dir(p), f"s{n}_contrast_{stamp()}"
                img0 = render.export(p, n, od / f"{tag}_before.png")
                changed = _fix_slide(p, n, rows_all, strat, out)
                mine = out["unfixable"][unfix_before:]
                after = [r for r in slide_contrast(p, n) if r["passes"] is False]
                left = [a for a in after if not any(u["slide"] == n and u["where"] == a["where"]
                                                    and u.get("run_start") == a.get("run_start")
                                                    and u.get("state") == a.get("state") for u in mine)]
                if left:
                    failures.append(f"still failing after the fix: {[(a['where'], a['text'], a['ratio']) for a in left][:5]}")
                vis = render.diff(img0, render.export(p, n, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
                for c in changed:
                    c["compare_image"] = vis.get("image")
            except Exception as exc:  # noqa: BLE001
                failures.append(f"{type(exc).__name__}: {com_err(exc)}")
            if failures:
                del out["unfixable"][unfix_before:]
                revert(p)
                out["failed"][n] = failures
                continue  # next strategy
            out["failed"].pop(n, None)
            if not changed:  # nothing could be changed (all unfixable): nothing to save
                if not p.Saved:
                    revert(p)
                break
            if autosave:
                save(p)
            out["changed"] += changed
            break
    return out
