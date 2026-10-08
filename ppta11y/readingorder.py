"""Reading order: what a screen reader reads on a slide, the order it should be in, and a verified way to
change it without changing how the slide looks or animates.

Screen readers, PowerPoint's Reading Order pane and tagged PDF export read a slide's top-level objects in
stacking order, back to front (slide.Shapes(1) first). Decorative and hidden objects are skipped, a group is
one stop (its alt text), and layout/master objects are not read at all. So the reading order IS the
stacking order, and fixing it means restacking. Restacking two objects is invisible when their painted
areas do not overlap. set_reading_order keeps the original stacking of every overlapping pair (so the
result is invisible by construction), then verifies with a before/after render, positions and the slide's
saved animation XML, and rolls back on any difference.

Suggested order (WCAG 1.3.2 Meaningful Sequence; Microsoft's guidance; matches PowerPoint's checker):
  1. the slide title, then a subtitle;
  2. everything else in visual reading order (rows top to bottom, columns left to right, recursively);
  3. footer, date and slide number last.
"""
from __future__ import annotations

import json
import math

from . import render
from .com import TITLE_PLACEHOLDERS, ToolError, com_err, safe
from .inventory import get_part, slide_title, text_of, timeline_targets
from .session import checkpoint, out_dir, require_editable, revert, save, stamp

FOOTER_PH = {13: "slide number", 14: "header", 15: "footer", 16: "date"}
SUBTITLE_PH = 4
ROW_TOL = 6.0          # points of overlap still treated as "separate rows/columns" ...
ROW_TOL_FRAC = 0.3     # ... or this fraction of the smaller object's height/width, whichever is larger
PAINT_MARGIN = 1.0     # points around every painted extent (antialiasing); the render check backs this up
EFFECT_MARGIN = 14.0   # extra points when a shape has a shadow / glow / soft edge
BLOCK_TAG, SEQ_TAG = "A11Y_BLOCK", "A11Y_SEQ"  # shapes our tools made that must be read together, in order


def _num(v, default=0.0) -> float:
    """A sane number from COM: PowerPoint returns -2147483648 ('mixed') or huge values for some effects."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return default
    return v if math.isfinite(v) and abs(v) < 100000 else default


# Entrance/exit effects that only mask the shape in place (appear, blinds, box, checkerboard, circle, diamond,
# dissolve, fade, plus, random bars, split, strips, wedge, wheel, wipe, color reveal). Anything else (fly,
# zoom, bounce, float, credits, grow, spin, motion paths, ...) can move/resize the shape mid-animation, so its
# painted extent is the whole slide.
STATIC_EFFECTS = {1, 3, 4, 5, 6, 8, 9, 10, 13, 14, 16, 18, 20, 21, 22, 27}
MOVING_BEHAVIORS = {1, 3, 4}       # motion, scale, rotation behaviours
MOVING_PROPERTIES = {1, 2, 3, 4, 6}  # property animations of x, y, width, height, rotation


# --------------------------------------------------------------------------------------------------
# Geometry
# --------------------------------------------------------------------------------------------------

def _rot_box(l, t, w, h, rot) -> tuple:
    if not rot:
        return (l, t, l + w, t + h)
    cx, cy, a = l + w / 2, t + h / 2, math.radians(rot)
    c, s = abs(math.cos(a)), abs(math.sin(a))
    hw, hh = (w * c + h * s) / 2, (w * s + h * c) / 2
    return (cx - hw, cy - hh, cx + hw, cy + hh)


def _union(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def _grow(b, m):
    return (b[0] - m, b[1] - m, b[2] + m, b[3] + m)


def overlaps(a, b, eps: float = 0.05) -> bool:
    if a is None or b is None:  # an object that paints nothing cannot hide or be hidden by anything
        return False
    return a[0] < b[2] - eps and b[0] < a[2] - eps and a[1] < b[3] - eps and b[1] < a[3] - eps


def _shape_box(s):
    return _rot_box(_num(s.Left), _num(s.Top), _num(s.Width), _num(s.Height), _num(safe(lambda: s.Rotation, 0)))


def _bounds(tr):
    v = [_num(safe(lambda: tr.BoundLeft), None), _num(safe(lambda: tr.BoundTop), None),
         _num(safe(lambda: tr.BoundWidth), None), _num(safe(lambda: tr.BoundHeight), None)]
    if None in v or v[2] <= 0 or v[3] <= 0:
        return None
    return (v[0], v[1], v[0] + v[2], v[1] + v[3])


def _rotate_rect(box, s):
    """Text bounds (BoundLeft/Top/Width/Height) of a rotated shape are already in slide space: PowerPoint
    reports the axis-aligned box of the rotated text (measured against the rendered ink at 0, 30, 90 and
    270 degrees). So nothing is rotated here; kept as the one place that would change if that ever differs."""
    return box


def _text_box(s, visible_only: bool = False):
    """Where the text is (it can overflow the shape; rotation applied), or None. visible_only: only
    paragraphs with visible characters (empty trailing lines otherwise stretch the box over things below)."""
    if not (safe(lambda: s.HasTextFrame) and safe(lambda: s.TextFrame.HasText)):
        return None
    tr = s.TextFrame.TextRange
    if visible_only:
        box = None
        n = safe(lambda: tr.Paragraphs().Count, 0) or 0
        for i in range(1, n + 1):
            p = tr.Paragraphs(i)
            if (safe(lambda: p.Text, "") or "").strip():
                box = _union(box, _bounds(p))
        if box is not None:
            return _rotate_rect(box, s)
    return _rotate_rect(_bounds(tr), s)


def _paints_box(s) -> bool:
    """True if the shape itself paints (fill, outline, picture, chart, table...), not only its text."""
    if s.Type in (13, 6, 3, 7, 10, 11, 16, 19, 24, 28, 29, 30, 31, 9, 5):  # pictures, groups, charts, tables...
        return True
    if s.Type == 14 and safe(lambda: s.PlaceholderFormat.ContainedType) not in (None, 1, 17, -2):
        return True  # a picture/chart/table/SmartArt held by a placeholder
    if any(safe(lambda k=k: getattr(s, k)) == -1 for k in ("HasChart", "HasTable", "HasSmartArt", "HasMedia")):
        return True
    fill = safe(lambda: s.Fill.Visible, 0) == -1 and (safe(lambda: s.Fill.Transparency, 0) or 0) < 0.999
    line = safe(lambda: s.Line.Visible, 0) == -1
    return bool(fill or line)


def _effects_margin(s) -> float:
    m = PAINT_MARGIN
    if safe(lambda: s.Line.Visible, 0) == -1:
        m += max(0.0, _num(safe(lambda: s.Line.Weight, 0)))
    if safe(lambda: s.Shadow.Visible, 0) == -1:
        m += EFFECT_MARGIN + max(abs(_num(safe(lambda: s.Shadow.OffsetX, 0))), abs(_num(safe(lambda: s.Shadow.OffsetY, 0)))) \
            + abs(_num(safe(lambda: s.Shadow.Blur, 0)))
    if _num(safe(lambda: s.Glow.Radius, 0)) > 0:
        m += EFFECT_MARGIN + _num(safe(lambda: s.Glow.Radius, 0))
    if _num(safe(lambda: s.SoftEdge.Radius, 0)) > 0:
        m += EFFECT_MARGIN
    return m


def _opaque_boxes(s) -> list:
    """Boxes the shape covers completely (solid, fully opaque fill, not rotated): what hides things beneath."""
    if safe(lambda: s.Visible, -1) == 0:
        return []
    if s.Type == 6:
        out = []
        for i in range(1, s.GroupItems.Count + 1):
            out += _opaque_boxes(s.GroupItems(i))
        return out
    if s.Type not in (1, 14, 17) or _num(safe(lambda: s.Rotation, 0)):
        return []
    if safe(lambda: s.AutoShapeType, 1) != 1:  # only plain rectangles cover their whole box
        return []
    f = safe(lambda: s.Fill)
    if f is None or safe(lambda: f.Visible, 0) != -1 or safe(lambda: f.Type) != 1 or _num(safe(lambda: f.Transparency, 0)) > 0:
        return []
    return [_shape_box(s)]


def _covered(box, covers: list, step: float = 3.0) -> bool:
    if box is None or not covers:
        return False
    x0, y0, x1, y1 = box
    nx, ny = max(2, int((x1 - x0) / step) + 1), max(2, int((y1 - y0) / step) + 1)
    for i in range(nx):
        x = x0 + (x1 - x0) * i / (nx - 1)
        for j in range(ny):
            y = y0 + (y1 - y0) * j / (ny - 1)
            if not any(c[0] <= x <= c[2] and c[1] <= y <= c[3] for c in covers):
                return False
    return True


def _clip(b, area):
    if b is None or area is None:
        return b
    c = (max(b[0], area[0]), max(b[1], area[1]), min(b[2], area[2]), min(b[3], area[3]))
    return c if c[0] < c[2] and c[1] < c[3] else None


def painted_box(s, slide_box=None, moving: bool = False):
    """Everything the shape can paint on the slide: rotated box (if it has a fill, outline or picture), its
    text (which can overflow the box), group children, outline and effects, clipped to the slide (nothing
    off the slide is ever shown). None if it paints nothing. A shape that moves or grows during an animation
    can paint anywhere on the slide."""
    if moving and slide_box is not None:
        return slide_box
    if s.Type == 6:
        box = None
        for i in range(1, s.GroupItems.Count + 1):
            box = _union(box, painted_box(s.GroupItems(i)))
    elif _paints_box(s):
        box = _union(_shape_box(s), _text_box(s))
    else:
        box = _text_box(s)
    if box is None:
        return None
    if safe(lambda: s.Reflection.Type, 0):
        box = (box[0], box[1], box[2], box[3] + (box[3] - box[1]))
    return _clip(_grow(box, _effects_margin(s)), slide_box)


def content_box(s):
    """Where a reader's eye finds the object: its text for anything holding text (a text box's frame can be
    far bigger than its words, and may have an invisible white fill), else the shape."""
    if s.Type not in (6, 13, 19, 3, 24, 28):
        tb = _text_box(s, visible_only=True)
        if tb is not None:
            return tb
    return _shape_box(s)


# --------------------------------------------------------------------------------------------------
# Items
# --------------------------------------------------------------------------------------------------

def _role(s) -> str:
    if s.Type != 14:
        return "object"
    ph = safe(lambda: s.PlaceholderFormat.Type)
    if ph in TITLE_PLACEHOLDERS:
        return "title"
    if ph == SUBTITLE_PH:
        return "subtitle"
    if ph in FOOTER_PH:
        return "footer"
    return "placeholder"


def _moving_ids(container) -> set[int]:
    """Shapes whose animation can move/enlarge them (their painted extent is the whole slide)."""
    from . import animations as anim
    out = set()
    for q in anim.snapshot(container):
        for e in q["effects"]:
            moves = any(b.get("Type") in MOVING_BEHAVIORS or
                        (b.get("Type") == 5 and (b.get("props") or {}).get("Property") in MOVING_PROPERTIES)
                        for b in e["behaviors"])
            if e["EffectType"] not in STATIC_EFFECTS or moves:
                out.add(e["shape_id"])
    return out


ARROW_AUTOSHAPES = set(range(33, 53)) | {89, 90, 91, 92}  # block arrows, curved arrows, chevrons, notched...


def _is_pointer(s) -> bool:
    """A line, connector or arrow: something that points at another object."""
    if s.Type == 9 or safe(lambda: s.Connector, 0) == -1:
        return True
    if safe(lambda: s.Line.Visible, 0) and (safe(lambda: s.Line.EndArrowheadStyle, 1) not in (1, -2)
                                            or safe(lambda: s.Line.BeginArrowheadStyle, 1) not in (1, -2)):
        return True
    return s.Type == 1 and safe(lambda: s.AutoShapeType) in ARROW_AUTOSHAPES


def _pointer_boxes(s) -> list:
    if s.Type != 6:
        return []
    return [_shape_box(s.GroupItems(i)) for i in range(1, s.GroupItems.Count + 1) if _is_pointer(s.GroupItems(i))]


def spoken(it: dict) -> str:
    k = it["kind"]
    if it["role"] == "title":
        k = "TITLE"
    elif it["role"] in ("subtitle", "footer"):
        k = it["role"] if it["role"] == "subtitle" else FOOTER_PH.get(it.get("ph"), "footer")
    if it["kind"] == "group" and it["alt"]:
        return f'group "{it["alt"][:70]}"'
    if it["text"]:
        return f'{k} "{it["text"][:70]}"'
    if it["alt"]:
        return f'{k} alt="{it["alt"][:70]}"'
    return f"{k} [nothing to read]"


def slide_items(pres, slide: int) -> list[dict]:
    """Top-level objects in stacking order (index 0 = read first), with what decides their reading order."""
    from .inventory import kind
    sl = get_part(pres, slide, "slide")
    W, H = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
    slide_box = (0.0, 0.0, W, H)
    targets, triggers = timeline_targets(sl)
    moving = _moving_ids(sl) if targets else set()
    morph = _morph(pres, slide)
    if morph:  # every shape may move during the transition
        moving = {x for i in range(1, sl.Shapes.Count + 1) for x in [sl.Shapes(i).Id] + (
            [sl.Shapes(i).GroupItems(k).Id for k in range(1, sl.Shapes(i).GroupItems.Count + 1)]
            if sl.Shapes(i).Type == 6 else [])}
    items = []
    opaque_above = []  # per index: opaque boxes of everything stacked above it (filled after the loop)
    shapes = [sl.Shapes(i) for i in range(1, sl.Shapes.Count + 1)]

    def _static_cover(s):
        """Only shapes that are always there can hide something: not animated (an exit or a click reveal
        uncovers what is underneath), not a trigger, visible, not an empty placeholder."""
        ids = {s.Id} | ({s.GroupItems(k).Id for k in range(1, s.GroupItems.Count + 1)} if s.Type == 6 else set())
        if ids & (set(targets) | set(triggers)) or not safe(lambda: s.Visible, -1):
            return []
        if s.Type == 14 and not text_of(s).strip() and safe(lambda: s.PlaceholderFormat.ContainedType) in (None, 1, 17, -2):
            return []
        return _opaque_boxes(s)
    covers = [_static_cover(s) for s in shapes]
    for i in range(len(shapes)):
        opaque_above.append([b for c in covers[i + 1:] for b in c])
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        leaves = [s.GroupItems(k).Id for k in range(1, s.GroupItems.Count + 1)] if s.Type == 6 else [s.Id]
        role = _role(s)
        txt = text_of(s).strip()
        dec = bool(safe(lambda: s.Decorative, 0))
        vis = bool(safe(lambda: s.Visible, -1))
        empty_ph = s.Type == 14 and not txt and safe(lambda: s.PlaceholderFormat.ContainedType) in (None, 1, 17, -2) \
            and not (s.AlternativeText or "").strip()
        anim_ids = [x for x in leaves + [s.Id] if targets.get(x) or x in triggers]
        items.append({
            "id": s.Id, "name": s.Name, "kind": kind(s), "role": role,
            "ph": safe(lambda: s.PlaceholderFormat.Type) if s.Type == 14 else None,
            "read": not dec and vis and not empty_ph,
            "decorative": dec, "hidden": not vis, "empty_placeholder": empty_ph,
            "alt": (s.AlternativeText or "").strip(), "text": txt.replace("\r", " / ").replace("\x0b", " ")[:160],
            "box": content_box(s),
            "paint": painted_box(s, slide_box, moving=bool(moving & set(leaves + [s.Id]))),
            "animated": bool(anim_ids), "moving": bool(moving & set(leaves + [s.Id])),
            "leaves": leaves, "pointers": _pointer_boxes(s),
            "area": max(1.0, _num(s.Width) * _num(s.Height)),
            "block": safe(lambda: s.Tags(BLOCK_TAG), "") or "",
            "seq": int(_num(safe(lambda: s.Tags(SEQ_TAG), 0) or 0)),
        })
        it = items[-1]
        # text that nobody can see: entirely behind opaque shapes stacked above it (e.g. a footer under a
        # full-slide table). It is still read; such text should be marked decorative.
        tb = _text_box(s) if not _paints_box(s) else None
        behind = bool(it["read"] and tb is not None and not it["animated"] and _covered(tb, opaque_above[i - 1]))
        # never for the title/subtitle: marking those decorative would take the slide's title away
        it["occluded"] = behind and role not in ("title", "subtitle")
        if behind and not it["occluded"]:
            it["title_behind_opaque"] = True
        if morph:
            it["morph"] = True
    return items


MORPH_EFFECTS = {3953, 3954, 3955}  # ppEffectMorphByObject / ByWord / ByChar


def _morph(pres, slide: int) -> bool:
    """Morph into or out of this slide moves matched shapes between their two positions, so any two shapes
    can cross mid-transition: their stacking must not change."""
    nxt = [slide + 1] if slide < pres.Slides.Count else []
    return any(safe(lambda n=n: pres.Slides(n).SlideShowTransition.EntryEffect) in MORPH_EFFECTS for n in [slide] + nxt)


# --------------------------------------------------------------------------------------------------
# Suggested order
# --------------------------------------------------------------------------------------------------

def _split(items: list[dict], axis: str) -> list[list[dict]]:
    lo, hi = (1, 3) if axis == "y" else (0, 2)
    s = sorted(items, key=lambda it: (it["box"][lo], it["box"][hi]))
    groups, cur = [], [s[0]]
    end, end_size = s[0]["box"][hi], s[0]["box"][hi] - s[0]["box"][lo]
    for it in s[1:]:
        size = it["box"][hi] - it["box"][lo]
        tol = max(ROW_TOL, ROW_TOL_FRAC * max(0.0, min(size, end_size)))
        if it["box"][lo] >= end - tol:
            groups.append(cur)
            cur, end, end_size = [it], it["box"][hi], size
        else:
            cur.append(it)
            if it["box"][hi] > end:
                end, end_size = it["box"][hi], size
    groups.append(cur)
    return groups


def _short_text(it) -> bool:
    return bool(it["text"]) and len(it["text"].split()) <= 12


def _prefer_columns(cols: list[list[dict]]) -> bool:
    """Columns that each start with a heading over bigger content, or a figure over a short caption, are
    read one column at a time (heading/figure, then what belongs to it), not row by row."""
    if len(cols) < 2:
        return False
    for c in cols:
        if len(c) < 2:
            return False
        a, b = sorted(c, key=lambda it: it["box"][1])[:2]
        heading = _short_text(a) and b["area"] >= 1.5 * a["area"]
        caption = (not a["text"] or a["kind"] in ("picture", "group", "chart", "table", "smartart")) and _short_text(b)
        if not (heading or caption):
            return False
    return True


def _colwise(items):
    return [x for c in _split(items, "x") for x in visual_order(c)]


def visual_order(items: list[dict]) -> list[dict]:
    """Recursive XY-cut: rows top to bottom, inside a row columns left to right, recursively. Heading/content
    and figure/caption columns are read column by column; a grid of such panels (2 rows per panel row) is
    read panel row by panel row, each panel row column by column. Objects that cannot be separated are read
    top to bottom, then left to right."""
    if len(items) <= 1:
        return list(items)
    rows = _split(items, "y")
    cols = _split(items, "x")
    if len(cols) > 1:
        if len(rows) == 1 or (len(rows) == 2 and _prefer_columns(cols)):
            return [x for c in cols for x in visual_order(c)]
        if len(rows) > 2 and len(rows) % 2 == 0:
            bands = [rows[i] + rows[i + 1] for i in range(0, len(rows), 2)]
            if all(len(_split(b, "x")) > 1 and _prefer_columns(_split(b, "x")) for b in bands):
                return [x for b in bands for x in _colwise(b)]
        if len(rows) > 2 and _prefer_columns(cols):
            # one heading row on top of several content rows: columns win only if no lower row is itself a
            # row of headings (that would be a grid of panels, read row by row)
            heads = [r for r in rows if len(r) == len(cols) and all(_short_text(it) for it in r)]
            if heads == [rows[0]]:
                return [x for c in cols for x in visual_order(c)]
    if len(rows) > 1:
        return [x for r in rows for x in visual_order(r)]
    return sorted(items, key=lambda it: (round(it["box"][1] / 4), it["box"][0]))


def _area(b) -> float:
    return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def _inter(a, b):
    return (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))


def annotations(rest: list[dict]) -> dict[int, int]:
    """{annotation id: target id}: a group whose arrow/line points into another object (a code note pointing
    at the code, a callout pointing at a figure) is read right after that object."""
    out = {}
    for a in rest:
        if not a["pointers"]:
            continue
        best, best_area = None, 0.0
        for t in rest:
            if t is a or t["pointers"]:
                continue
            ar = sum(_area(_inter(_grow(p, 4), t["box"])) for p in a["pointers"])
            if ar > best_area:
                best, best_area = t, ar
        if best is not None and best_area > 0:
            out[a["id"]] = best["id"]
    return out


def _place_annotations(ordered: list[dict]) -> list[dict]:
    notes = annotations(ordered)
    if not notes:
        return ordered
    base = [it for it in ordered if it["id"] not in notes]
    out = []
    for it in base:
        out.append(it)
        out += [n for n in ordered if notes.get(n["id"]) == it["id"]]  # keeps the notes' own visual order
    return out


def _units(items: list[dict]) -> list[dict]:
    """Shapes tagged as one block (e.g. a table converted to text boxes, already in its reading order) are
    ordered as a single unit and expanded back in their own sequence."""
    units, blocks = [], {}
    for it in items:
        if it.get("block"):
            blocks.setdefault(it["block"], []).append(it)
        else:
            units.append(it)
    for name, mem in blocks.items():
        mem = sorted(mem, key=lambda x: x["seq"])
        box = mem[0]["box"]
        for m in mem[1:]:
            box = _union(box, m["box"])
        units.append({**mem[0], "id": f"block:{name}", "box": box, "area": sum(m["area"] for m in mem),
                      "pointers": [], "members": mem, "kind": "block"})
    return units


def _expand(units: list[dict]) -> list[dict]:
    out = []
    for u in units:
        out += u.get("members", [u])
    return out


def suggest(items: list[dict], variant: str = "rules") -> list[dict]:
    """The order a screen reader should use. variant:
      rules         title, subtitle, content in visual order with call-out notes right after what they
                    point at, footers last (default)
      visual        the same without moving call-out notes (pure visual order)
      placeholders  title, subtitle, other text placeholders, then other objects, footers last
      top_down      title, then strictly by top edge, footers last
      minimal       the current order with only the title moved first and footers moved last
    """
    read = [it for it in items if it["read"]]
    title = [it for it in read if it["role"] == "title"]
    sub = [it for it in read if it["role"] == "subtitle"]
    foot = sorted([it for it in read if it["role"] == "footer"], key=lambda it: (round(it["box"][1] / 4), it["box"][0]))
    rest = [it for it in read if it["role"] not in ("title", "subtitle", "footer")]
    if variant == "minimal":
        return title + [it for it in read if it["role"] not in ("title", "footer")] + foot
    units = _units(rest)
    if variant == "placeholders":
        ph = [u for u in units if u["role"] == "placeholder"]
        return title + sub + _expand(visual_order(ph)) + _expand(_place_annotations(
            visual_order([u for u in units if u["role"] != "placeholder"]))) + foot
    if variant == "top_down":
        return title + sub + _expand(_place_annotations(sorted(units, key=lambda u: (u["box"][1], u["box"][0])))) + foot
    if variant == "visual":
        return title + sub + _expand(visual_order(units)) + foot
    return title + sub + _expand(_place_annotations(visual_order(units))) + foot


VARIANTS = ("rules", "visual", "placeholders", "top_down", "minimal")


def issues(items: list[dict]) -> list[str]:
    read = [it for it in items if it["read"]]
    out = []
    if not read:
        return out
    ti = next((i for i, it in enumerate(read) if it["role"] == "title"), None)
    if ti is None:
        out.append("no slide title is read (add one: titles tools)")
    elif ti > 0:
        out.append(f"title is read in position {ti + 1}, after: " + "; ".join(spoken(it) for it in read[:ti]))
    fi = next((i for i, it in enumerate(read) if it["role"] == "footer"), None)
    if fi is not None and any(it["role"] != "footer" for it in read[fi + 1:]):
        out.append("footer/slide number is read before slide content: " +
                   "; ".join(spoken(it) for it in read[fi + 1:] if it["role"] != "footer"))
    hid = [it for it in read if it.get("occluded")]
    if hid:
        out.append("read but invisible (completely behind other objects), mark decorative: " +
                   "; ".join(f"#{it['id']} {spoken(it)}" for it in hid))
    return out


# --------------------------------------------------------------------------------------------------
# Accepted orders (an order that was verified and applied, e.g. chosen by the agent for meaning, counts as
# done even when it differs from the rules suggestion). Keyed by SlideID, which survives slide moves.
# --------------------------------------------------------------------------------------------------

def _accepted_path(pres):
    return out_dir(pres) / "reading_order_accepted.json"


def load_accepted(pres) -> dict:
    try:
        return json.loads(_accepted_path(pres).read_text(encoding="utf8"))
    except FileNotFoundError:
        return {}
    except Exception as exc:
        raise ToolError(f"cannot read {_accepted_path(pres)}: {exc}")


def _record_accepted(pres, slide: int, order: list[int], how: str) -> None:
    import os
    p = _accepted_path(pres)
    m = load_accepted(pres)
    m[str(pres.Slides(slide).SlideID)] = {"order": order, "how": how, "slide_at_time": slide}
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(m, indent=1), encoding="utf8")
    os.replace(tmp, p)


# --------------------------------------------------------------------------------------------------
# Reports
# --------------------------------------------------------------------------------------------------

def slide_report(pres, slide: int, variant: str = "rules", accepted: dict | None = None) -> dict:
    items = slide_items(pres, slide)
    read = [it for it in items if it["read"]]
    sug = suggest(items, variant)
    now_ids, sug_ids = [it["id"] for it in read], [it["id"] for it in sug]
    sl = pres.Slides(slide)
    acc = (accepted if accepted is not None else load_accepted(pres)).get(str(sl.SlideID))
    accepted_now = bool(acc) and acc.get("order") == now_ids
    notes = [f"the title #{it['id']} is behind opaque objects (not visible, still read as the title; kept)"
             for it in read if it.get("title_behind_opaque")]
    if items and items[0].get("morph"):
        notes.append("Morph transition into/out of this slide: stacking is kept as it is (any restack could change "
                     "the transition), so only the order already possible is offered")
    return {
        **({"notes": notes} if notes else {}),
        "slide": slide, "title": slide_title(sl), "layout": sl.CustomLayout.Name,
        "reads_now": [f"{i}. #{it['id']} {spoken(it)}" for i, it in enumerate(read, 1)],
        "order_now": now_ids,
        "suggested": sug_ids,
        "suggested_reads": [f"{i}. #{it['id']} {spoken(it)}" for i, it in enumerate(sug, 1)],
        "matches_suggestion": (now_ids == sug_ids or accepted_now) and not any(it.get("occluded") for it in read),
        **({"accepted_order": acc.get("how")} if accepted_now and now_ids != sug_ids else {}),
        "issues": issues(items),
        "not_read": [f"#{it['id']} {it['kind']}" + (" (decorative)" if it["decorative"] else " (hidden)" if it["hidden"]
                                                     else " (empty placeholder)") for it in items if not it["read"]],
    }


def _original_leaf_order(src_pres, slide: int) -> list[int]:
    sl = src_pres.Slides(slide)
    out = []
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        out += [s.GroupItems(k).Id for k in range(1, s.GroupItems.Count + 1)] if s.Type == 6 else [s.Id]
    return out


def deck_report(pres, slides: list[int] | None = None, compare_to: str | None = None, write: bool = True) -> dict:
    """Reading order of every slide: what is read now, the suggested order, issues; optionally what our
    edits changed compared with the original deck (`compare_to`, opened read-only)."""
    from . import session
    slides = slides or list(range(1, pres.Slides.Count + 1))
    src, opened = None, False
    if compare_to:
        src = session.find_open(compare_to)
        if src is None:
            from .com import TRUE, app
            src = app().Presentations.Open(compare_to, TRUE, 0, 0)
            opened = True
    rows = []
    accepted = load_accepted(pres)
    try:
        for n in slides:
            r = slide_report(pres, n, accepted=accepted)
            if src is not None and n <= src.Slides.Count:
                pos = {lid: i for i, lid in enumerate(_original_leaf_order(src, n))}
                items = {it["id"]: it for it in slide_items(pres, n)}
                firsts = []
                for sid in r["order_now"]:
                    p = [pos[x] for x in items[sid]["leaves"] if x in pos]
                    firsts.append((sid, min(p) if p else None))
                known = [(sid, p) for sid, p in firsts if p is not None]
                moved = [f"#{a} now read before #{b}" for i, (a, pa) in enumerate(known) for b, pb in known[i + 1:] if pa > pb]
                r["changed_vs_original"] = moved[:12]
                r["new_vs_original"] = [sid for sid, p in firsts if p is None]
            rows.append(r)
    finally:
        if opened:
            safe(lambda: src.Close())
    summary = {
        "slides_checked": len(rows),
        "slides_matching_suggestion": sum(r["matches_suggestion"] for r in rows),
        "slides_to_fix": [r["slide"] for r in rows if not r["matches_suggestion"]],
        "title_not_first": [r["slide"] for r in rows if any(i.startswith("title is read") for i in r["issues"])],
        "footer_before_content": [r["slide"] for r in rows if any(i.startswith("footer") for i in r["issues"])],
        "no_title": [r["slide"] for r in rows if any(i.startswith("no slide title") for i in r["issues"])],
    }
    out = {"deck": pres.FullName, "summary": summary, "slides": rows}
    if write:
        d = out_dir(pres)
        base = d / f"reading_order_{stamp()}"
        base.with_suffix(".json").write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf8")
        lines = []
        for r in rows:
            lines.append(f"Slide {r['slide']} - {r['title'] or '(no title)'}  [{r['layout']}]"
                         + ("" if r["matches_suggestion"] else "   <-- order should change"))
            lines += ["   now:       " + x for x in r["reads_now"]]
            if not r["matches_suggestion"]:
                lines += ["   suggested: " + x for x in r["suggested_reads"]]
            lines += ["   !! " + x for x in r["issues"]]
            if r.get("changed_vs_original"):
                lines.append("   vs original: " + "; ".join(r["changed_vs_original"]))
            lines.append("")
        base.with_suffix(".txt").write_text("\n".join(lines), encoding="utf8")
        out["report_files"] = {"json": str(base.with_suffix(".json")), "txt": str(base.with_suffix(".txt"))}
    return out


# --------------------------------------------------------------------------------------------------
# Restacking
# --------------------------------------------------------------------------------------------------

def _slot_order(current: list[int], read_set: set[int], order: list[int]) -> list[int]:
    """Put the read objects into the slots read objects occupy now, in the requested order; everything
    else (decorative, hidden) keeps its slot."""
    slots = [i for i, sid in enumerate(current) if sid in read_set]
    target = list(current)
    for i, sid in zip(slots, order):
        target[i] = sid
    return target


def _constrained_order(items: list[dict], order: list[int], relax: set | None = None) -> list[int]:
    """A full stacking order that keeps the current relative order of every pair of objects whose painted
    areas overlap (so nothing can look different) and otherwise follows `order` as closely as possible.
    `relax` = pairs (a, b) allowed to swap (they must then be proven invisible by a render)."""
    relax = relax or set()
    cur = [it["id"] for it in items]
    pos = {sid: i for i, sid in enumerate(cur)}
    rank = {sid: i for i, sid in enumerate(order)}
    by = {it["id"]: it for it in items}
    prio = {}
    for sid in cur:
        if sid in rank:
            prio[sid] = (rank[sid], 0, pos[sid])
        else:  # keep non-read objects just before the read object that followed them
            nxt = next((rank[x] for x in cur[pos[sid] + 1:] if x in rank), len(order))
            prio[sid] = (nxt, -1, pos[sid])
    preds = {sid: set() for sid in cur}
    for i, a in enumerate(cur):
        for b in cur[i + 1:]:
            if overlaps(by[a]["paint"], by[b]["paint"]) and (a, b) not in relax and (b, a) not in relax:
                preds[b].add(a)  # a is painted under b now; keep it that way
    out, done = [], set()
    while len(out) < len(cur):
        ready = [sid for sid in cur if sid not in done and preds[sid] <= done]
        nxt = min(ready, key=lambda s: prio[s])
        out.append(nxt)
        done.add(nxt)
    return out


def _flipped_overlaps(items: list[dict], target: list[int]) -> list[tuple[int, int]]:
    cur = [it["id"] for it in items]
    pc, pt = {s: i for i, s in enumerate(cur)}, {s: i for i, s in enumerate(target)}
    by = {it["id"]: it for it in items}
    out = []
    for i, a in enumerate(cur):
        for b in cur[i + 1:]:
            if (pt[a] < pt[b]) != (pc[a] < pc[b]) and overlaps(by[a]["paint"], by[b]["paint"]):
                out.append((a, b))
    return out


RISKY_THRESHOLD = 2   # a restack of overlapping objects must not change any channel by more than this
RISKY_EDGE_PX = 400   # ... except a thin line of edge pixels that a 1-px antialiasing shift explains


def _animation_near(items: list[dict], by: dict, a: int, b: int) -> bool:
    """Can the still render miss a state in which the stacking of a and b shows? Yes if either moves,
    grows or spins during an animation, or if another animated object (or trigger) is painted over their
    overlap (an exit or a click reveal uncovers it). An appear/fade/wipe-type effect on a or b itself is
    fine: when the render is identical their painted pixels do not conflict, in any frame."""
    if by[a]["moving"] or by[b]["moving"]:
        return True
    box = _inter(by[a]["paint"], by[b]["paint"])
    return any(it["animated"] and it["id"] not in (a, b) and overlaps(it["paint"], box) for it in items)


def _must_swap(items: list[dict], order: list[int]) -> set:
    """Pairs of READ objects that overlap and are stacked against the requested order: the only pairs the
    requested order needs to swap (everything else, decorative objects included, keeps its stacking)."""
    rank = {sid: i for i, sid in enumerate(order)}
    cur = [it for it in items if it["id"] in rank]
    out = set()
    for i, a in enumerate(cur):
        for b in cur[i + 1:]:
            if rank[a["id"]] > rank[b["id"]] and overlaps(a["paint"], b["paint"]):
                out.add((a["id"], b["id"]))
    return out


def _movers(have: list[int], want: list[int]) -> set:
    """The fewest objects to move so `have` reads as `want`: everything outside their longest common
    subsequence (e.g. only the title, when it is read last instead of first)."""
    n, m = len(have), len(want)
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            dp[i][j] = dp[i + 1][j + 1] + 1 if have[i] == want[j] else max(dp[i + 1][j], dp[i][j + 1])
    keep, i, j = set(), 0, 0
    while i < n and j < m:
        if have[i] == want[j]:
            keep.add(want[j])
            i, j = i + 1, j + 1
        elif dp[i + 1][j] >= dp[i][j + 1]:
            i += 1
        else:
            j += 1
    return set(want) - keep


def _restack(sl, target: list[int]) -> None:
    by = {sl.Shapes(i).Id: sl.Shapes(i) for i in range(1, sl.Shapes.Count + 1)}
    cur = [sl.Shapes(i).Id for i in range(1, sl.Shapes.Count + 1)]
    if cur == target:
        return
    for sid in target:  # bring each to the front in turn: the last one ends on top
        by[sid].ZOrder(0)  # msoBringToFront
    now = [sl.Shapes(i).Id for i in range(1, sl.Shapes.Count + 1)]
    if now != target:
        raise ToolError(f"PowerPoint did not apply the stacking order (got {now}, wanted {target})")


def set_reading_order(pres, slide: int, order: list[int], autosave: bool = True, strict: bool = False,
                      how: str = "agent") -> dict:
    """Make the screen-reader order of `slide` equal `order` (every non-decorative, visible top-level
    object id exactly once) without any visible change. Objects whose painted areas overlap keep their
    relative stacking; if that prevents part of the requested order, the closest invisible order is used
    and the difference reported (strict=True fails instead). Verified: render, positions, animation XML."""
    from .ops import VISUAL_TOLERANCE_PX, _geometry, _geom_diffs, _timing_now, _xml_compare
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    items = slide_items(pres, slide)
    read_ids = [it["id"] for it in items if it["read"]]
    order = [int(x) for x in order]
    if len(set(order)) != len(order):
        raise ToolError("order lists an id twice")
    missing = [i for i in read_ids if i not in order]
    extra = [i for i in order if i not in read_ids]
    if missing or extra:
        raise ToolError(f"order must list every object a screen reader reads on slide {slide} exactly once: "
                        f"missing {missing}, not readable/not on the slide {extra}. Readable now: {read_ids}")
    current = [it["id"] for it in items]
    by = {it["id"]: it for it in items}
    if read_ids == order:
        _record_accepted(pres, slide, order, how)
        return {"slide": slide, "ok": True, "unchanged": True, "reads_now": [spoken(by[i]) for i in order]}

    target = _slot_order(current, set(read_ids), order)
    flips = _flipped_overlaps(items, target)
    plan = "invisible by construction"
    if flips:
        alt = _constrained_order(items, order)
        alt_reads = [s for s in alt if s in set(read_ids)]
        swap = _must_swap(items, order)
        trial = _constrained_order(items, order, relax=swap) if swap else alt
        if [x for x in trial if x in set(read_ids)] != order:
            # a decorative/hidden object (e.g. a converted table's border lines) touches an object that must
            # move: let only the misplaced objects swap with what they touch; the render still has to match
            mv = _movers(alt_reads, order)
            swap = swap | {(m_, it["id"]) for m_ in mv for it in items
                           if it["id"] != m_ and overlaps(by[m_]["paint"], it["paint"])}
            trial = _constrained_order(items, order, relax=swap)
        if alt_reads == order:
            target, plan = alt, "invisible by construction (overlapping objects keep their stacking)"
        elif strict:
            raise ToolError(f"the requested order would restack overlapping objects {sorted(swap)}; closest invisible "
                            f"order reads {alt_reads}. Use that order or strict=False")
        elif any(_animation_near(items, by, a, b) for a, b in swap):
            # an object that moves, or another animated object over their overlap, can show a stacking the
            # still render never shows; keep such pairs
            target, plan = alt, "closest invisible order (overlapping objects near an animation keep their stacking)"
        elif [x for x in trial if x in set(read_ids)] != order:
            target, plan = alt, "closest invisible order (decorative or hidden objects in between keep their stacking)"
        else:
            # only the read objects that must swap are swapped; boxes overlap but the pixels may not: keep it
            # only if the render is identical
            target, plan = trial, "requested order; overlapping pairs restacked, accepted only if the render is identical"

    checkpoint(pres)
    slide_id = pres.Slides(slide).SlideID
    has_anim = bool(timeline_targets(sl)[0])
    result = {"slide": slide, "reads_before": [spoken(by[i]) for i in read_ids], "plan": plan}
    failures: list[str] = []
    try:
        od = out_dir(pres)
        tag = f"s{slide}_ro_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        leaves = {x for it in items for x in it["leaves"]}
        geom0 = _geometry(sl, leaves)
        xml0 = _timing_now(pres, slide_id) if has_anim else None
        _restack(sl, target)
        risky = plan.startswith("requested order; overlapping")
        # overlapping pairs were restacked on purpose: accept only an identical render (any channel changing
        # by more than 2 levels counts), or one whose few differing pixels are all explained by a 1-px
        # antialiasing shift (an edge touching an edge)
        vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png",
                          threshold=RISKY_THRESHOLD if risky else 24)
        if risky and vis["changed_pixels"] > 0:
            soft = (render.tolerant_changed(img0, od / f"{tag}_after.png", threshold=RISKY_THRESHOLD)
                    if vis["changed_pixels"] <= RISKY_EDGE_PX else 1)
            vis["shift_tolerant_changed"] = soft
        if risky and vis["changed_pixels"] > 0 and vis.get("shift_tolerant_changed", 1) > 0:
            alt = _constrained_order(items, order)
            _restack(sl, alt)
            target, plan = alt, "closest invisible order (restacking the overlapping objects was visible)"
            vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after2.png"), od / f"{tag}_compare.png")
            if strict:
                failures.append("the requested order is only possible with a visible change")
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"slide looks different after restacking ({vis['changed_pixels']} px, see {vis.get('image')})")
        failures += _geom_diffs(geom0, _geometry(sl, leaves))
        xml = None
        if xml0 is not None:
            xml = _xml_compare(xml0, _timing_now(pres, slide_id))
            if xml["status"] not in ("identical", "equivalent"):
                failures.append(f"animation XML changed ({xml['status']})")
        now = slide_items(pres, slide)
        reads = [it["id"] for it in now if it["read"]]
        left = issues(now)
        result.update(plan=plan, order_requested=order, order_now=reads,
                      reads_now=[spoken(by[i]) for i in reads], visual_check=vis,
                      xml_check=(xml or {"status": "no animations"}).get("status"))
        if reads != order:
            result["deviation"] = ("overlapping objects must keep their stacking, so this is the closest order that "
                                   "looks identical")
        if left:
            result["issues_left"] = left
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    if failures:
        revert(pres)
        result.update(ok=False, reverted=True, failures=failures)
        return result
    if autosave:
        save(pres)
    # done = an order the agent chose for meaning, or one that meets the rules (title first, footers last).
    # A closest-invisible order was chosen by the constraint, not by anyone: it needs a look first.
    if result.get("deviation"):
        result["not_accepted"] = ("saved, but this is the closest order that looks identical, not the one asked for. "
                                  "Check reads_now; if it makes sense, accept it with set_reading_order(slide, "
                                  f"{result['order_now']}) (exactly this order)" +
                                  ("; also: " + "; ".join(result["issues_left"]) if result.get("issues_left") else ""))
    elif how == "agent" or not result.get("issues_left"):
        _record_accepted(pres, slide, result["order_now"], how)
    else:
        result["not_accepted"] = "saved, but the slide still needs attention: " + "; ".join(result["issues_left"])
    result.update(ok=True, saved=autosave)
    return result


def mark_hidden_text_decorative(pres, slide: int) -> list[str]:
    """Text completely behind static opaque objects (nobody can see it) is marked decorative so screen
    readers do not read it either. Each one is proven invisible first: hiding it must not change a single
    pixel of the slide. Saved. Returns what was marked."""
    from .ops import mark_decorative
    from .inventory import find_shape
    hid = [it for it in slide_items(pres, slide) if it.get("occluded")]
    if not hid:
        return []
    checkpoint(pres)
    marked = []
    try:
        sl = get_part(pres, slide, "slide")
        od, tag = out_dir(pres), f"s{slide}_hid_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        for it in hid:
            s = find_shape(sl, it["id"])
            s.Visible = 0
            changed = render.diff(img0, render.export(pres, slide, od / f"{tag}_{it['id']}.png"))["changed_pixels"]
            s.Visible = -1
            if changed == 0:
                marked.append(it)
        if marked:
            mark_decorative(pres, slide, [it["id"] for it in marked])
        if render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"))["changed_pixels"]:
            raise ToolError("slide changed while checking hidden text")
        save(pres)
    except Exception:
        revert(pres)
        raise
    return [f"#{it['id']} {spoken(it)}" for it in marked]


def fix_reading_order(pres, slides: list[int] | None = None, orders: dict | None = None, use_checker: bool = True,
                      variants: tuple = VARIANTS, hidden_text: str = "decorative") -> dict:
    """Apply the suggested order (or `orders` = {slide: [ids]} chosen by the agent) to every slide, each as its
    own verified transaction. Text nobody can see (completely behind opaque objects) is first marked
    decorative (hidden_text="decorative"; "keep" leaves it). With use_checker, PowerPoint's Accessibility
    Assistant is then run; for slides it still flags under "Check reading order" the other suggestion variants
    are tried (each invisible and verified) and the first one PowerPoint accepts is kept."""
    from . import checker, session
    path = pres.FullName

    def live():  # a failed transaction reopens the deck from disk: always use the current object
        p = session.find_open(path)
        if p is None:
            raise ToolError(f"{path} is no longer open")
        return p

    slides = slides or list(range(1, pres.Slides.Count + 1))
    try:
        orders = {int(k): [int(x) for x in v] for k, v in (orders or {}).items()}
    except (TypeError, ValueError):
        raise ToolError("orders must be {slide: [shape ids]} with plain integer ids ('@N' ids are resolved by "
                        "apply_plan, not here)")
    out = {"applied": {}, "unchanged": [], "failed": {}, "checker": None, "hidden_text_marked_decorative": {},
           "hidden_text_failed": {}, "needs_attention": {}}
    chosen = {}
    if hidden_text == "decorative":
        for n in slides:
            try:
                done = mark_hidden_text_decorative(live(), n)
            except Exception as exc:  # noqa: BLE001 - recorded; that slide keeps its hidden text
                out["hidden_text_failed"][n] = f"{type(exc).__name__}: {com_err(exc)}"
                continue
            if done:
                out["hidden_text_marked_decorative"][n] = done

    def apply(n: int, order, label: str):
        """`order` is a list of ids or a function of the live deck giving one (computed here, inside the
        try, so one slide's error never stops the others)."""
        try:
            p = live()
            if callable(order):
                order = order(p)
            readable = {it["id"] for it in slide_items(p, n) if it["read"]}
            order = [i for i in order if i in readable]  # e.g. hidden text just marked decorative
            res = set_reading_order(p, n, order, how=label)
        except Exception as exc:  # noqa: BLE001 - one slide must not stop the others
            out["failed"][n] = [f"{type(exc).__name__}: {com_err(exc)}"]
            return False
        if not res.get("ok"):
            out["failed"][n] = res.get("failures")
            return False
        out["failed"].pop(n, None)
        if res.get("not_accepted"):
            out["needs_attention"][n] = res["not_accepted"]
        else:
            out["needs_attention"].pop(n, None)
        if res.get("unchanged"):
            if n not in out["unchanged"]:
                out["unchanged"].append(n)
        else:
            out["applied"][n] = {"variant": label, "reads_now": res.get("reads_now"), "plan": res.get("plan"),
                                 "order_now": res.get("order_now"),
                                 **({"deviation": res["deviation"]} if res.get("deviation") else {})}
        return True

    def by_variant(n: int, v: str):
        return lambda p: [it["id"] for it in suggest(slide_items(p, n), v)]

    def flags():
        try:
            return _checker_reading_flags(live(), checker)
        except Exception as exc:  # noqa: BLE001 - the checker UI can fail; the rules orders stay applied
            out["checker"] = {"error": f"{type(exc).__name__}: {com_err(exc)}"}
            return None

    for n in slides:
        if n in orders:
            apply(n, orders[n], "agent")
            chosen[n] = "agent"
        else:
            apply(n, by_variant(n, "rules"), "rules")
            chosen[n] = "rules"
    if use_checker:
        flagged = flags()
        tried: dict[int, set] = {}
        for v in variants:
            if flagged is None:
                break
            todo = [n for n in flagged if n in slides and n not in orders
                    and v not in tried.setdefault(n, {chosen.get(n, "rules")})]
            if not todo:
                continue
            for n in todo:
                tried[n].add(v)
                try:
                    cand = suggest(slide_items(live(), n), v)
                except Exception as exc:  # noqa: BLE001
                    out["failed"][n] = [f"variant {v}: {type(exc).__name__}: {com_err(exc)}"]
                    continue
                if issues([dict(it, read=True) for it in cand]):  # never trade the rules away for the checker
                    continue
                apply(n, [it["id"] for it in cand], v)
            flagged = flags()
            for n in todo:
                if flagged is not None and n not in flagged:
                    chosen[n] = v
        if flagged is not None:
            # anything still flagged: go back to the rules order (the most meaningful one)
            for n in [n for n in flagged if n in slides and n not in orders]:
                apply(n, by_variant(n, "rules"), "rules (PowerPoint still flags)")
            flagged = flags()
        if flagged is not None:
            out["checker"] = {"still_flagged": flagged,
                              "note": ("PowerPoint's 'Check reading order' is a heuristic prompt to verify; slides listed "
                                       "here were checked against the rules (title first, visual order, footers last)")
                              if flagged else "PowerPoint's checker reports no reading-order warnings"}
    out["variant_used"] = chosen
    return out


def _checker_reading_flags(pres, checker) -> list[int]:
    res = checker.run(pres, itemize=["Check reading order"])
    return sorted({it.get("slide") for it in res["items"].get("Check reading order", []) if it.get("slide")})
