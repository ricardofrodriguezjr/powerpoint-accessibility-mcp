"""Read-only views of slides, layouts and masters: shapes, alt text, decorative flags, animation targets."""
from __future__ import annotations

from collections import Counter

from .com import NEEDS_ALT_TYPES, SHAPE_TYPES, TEXT_PLACEHOLDERS, TITLE_PLACEHOLDERS, ToolError, safe

PARTS = ("slide", "layout", "master")


def get_part(pres, slide: int, part: str = "slide"):
    if part not in PARTS:
        raise ToolError(f"part must be one of {PARTS}")
    if not 1 <= slide <= pres.Slides.Count:
        raise ToolError(f"slide {slide} out of range 1..{pres.Slides.Count}")
    sl = pres.Slides(slide)
    if part == "slide":
        return sl
    if part == "layout":
        return sl.CustomLayout
    return sl.Master


def part_label(pres, slide: int, part: str) -> str:
    if part == "slide":
        return f"slide {slide}"
    if part == "layout":
        lay = pres.Slides(slide).CustomLayout
        users = slides_using_layout(pres, lay)
        return f"layout '{lay.Name}' (used by slides {_ranges(users)})"
    return f"slide master '{pres.Slides(slide).Master.Name}'"


def layout_key(lay) -> tuple:
    """Identity of a layout (names can repeat across masters)."""
    return (lay.Design.Index, lay.Index)


def slides_using_layout(pres, lay, shown_only: bool = False) -> list[int]:
    """Slides that use `lay`; shown_only=True keeps only slides that display its graphics
    (not 'Hide background graphics')."""
    key, out = layout_key(lay), []
    for i in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(i)
        if layout_key(sl.CustomLayout) == key and (not shown_only or safe(lambda: sl.DisplayMasterShapes, -1)):
            out.append(i)
    return out


def slides_showing_master(pres, design_index: int) -> list[int]:
    out = []
    for i in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(i)
        if sl.Design.Index == design_index and safe(lambda: sl.DisplayMasterShapes, -1)                 and safe(lambda: sl.CustomLayout.DisplayMasterShapes, -1):
            out.append(i)
    return out


def _ranges(nums: list[int]) -> str:
    if not nums:
        return "(none)"
    out, start, prev = [], nums[0], nums[0]
    for n in nums[1:] + [None]:
        if n is not None and n == prev + 1:
            prev = n
            continue
        out.append(f"{start}" if start == prev else f"{start}-{prev}")
        if n is not None:
            start = prev = n
    return ", ".join(out)


def iter_shapes(shapes, depth: int = 0, parent=None):
    """Yield (shape, depth, parent_group) for every shape, recursing into groups."""
    for i in range(1, shapes.Count + 1):
        s = shapes(i)
        yield s, depth, parent
        if s.Type == 6:
            yield from _iter_group(s, depth + 1)


def _iter_group(g, depth):
    items = g.GroupItems
    for i in range(1, items.Count + 1):
        s = items(i)
        yield s, depth, g
        if s.Type == 6:  # nested group (GroupItems is flat in COM, but be defensive)
            yield from _iter_group(s, depth + 1)


def find_shape(container, sid: int):
    for s, _, _ in iter_shapes(container.Shapes):
        if s.Id == sid:
            return s
    raise ToolError(f"no shape with id {sid} on this part")


def top_level_index(container, sid: int) -> int | None:
    shapes = container.Shapes
    for i in range(1, shapes.Count + 1):
        if shapes(i).Id == sid:
            return i
    return None


def timeline_targets(container) -> tuple[Counter, set]:
    """(effects per target shape id, ids used as interactive triggers)."""
    targets, triggers = Counter(), set()
    tl = safe(lambda: container.TimeLine)
    if tl is None:
        return targets, triggers
    seqs = [tl.MainSequence] + [tl.InteractiveSequences(k) for k in range(1, tl.InteractiveSequences.Count + 1)]
    for q in seqs:
        for i in range(1, q.Count + 1):
            e = q.Item(i)
            sid = safe(lambda: e.Shape.Id)
            if sid is not None:
                targets[sid] += 1
            trig = safe(lambda: e.Timing.TriggerShape.Id)
            if trig is not None:
                triggers.add(trig)
    return targets, triggers


def decorative(s):
    v = safe(lambda: s.Decorative)
    return None if v is None else bool(v)


def text_of(s) -> str:
    if safe(lambda: s.HasTextFrame) and safe(lambda: s.TextFrame.HasText):
        return safe(lambda: s.TextFrame.TextRange.Text, "") or ""
    if s.Type == 19 or safe(lambda: s.HasTable):
        return "[table]"
    return ""


def placeholder_info(s) -> dict | None:
    if s.Type != 14:
        return None
    pf = safe(lambda: s.PlaceholderFormat)
    if pf is None:
        return None
    return {"type": safe(lambda: pf.Type), "contains": SHAPE_TYPES.get(safe(lambda: pf.ContainedType), None)}


def kind(s) -> str:
    """Human label; placeholders are reported by what they contain (a picture in a content placeholder
    is a picture for accessibility purposes)."""
    if s.Type == 14:
        ph = placeholder_info(s) or {}
        c = ph.get("contains")
        if c and c not in ("autoshape", "textbox"):
            return f"{c} (in placeholder)"
        return "text placeholder" if ph.get("type") in TEXT_PLACEHOLDERS else "placeholder"
    return SHAPE_TYPES.get(s.Type, f"type{s.Type}")


def effective_type(s) -> int:
    if s.Type == 14:
        ct = safe(lambda: s.PlaceholderFormat.ContainedType)
        if ct not in (None, 1, 17, -2):
            return ct
    return s.Type


def needs_alt(s) -> str:
    """'required' (checker flags it), 'review' (a shape with no text: alt text or decorative),
    'no' (text is read directly; tables are read cell by cell)."""
    t = effective_type(s)
    if t in NEEDS_ALT_TYPES:
        return "required"
    if t in (1, 2, 5, 15, 17) and not text_of(s).strip():  # incl. text boxes holding only spaces
        return "review"
    if t == 9:  # lines and connectors: the checker usually flags them too
        return "review"
    return "no"


def shape_info(s, targets: Counter | None = None, triggers: set | None = None, deep: bool = True) -> dict:
    targets = targets or Counter()
    d = {
        "id": s.Id, "name": s.Name, "kind": kind(s), "type": s.Type,
        "alt": s.AlternativeText or "", "title": safe(lambda: s.Title, "") or "",
        "decorative": decorative(s),
        "left": round(s.Left, 1), "top": round(s.Top, 1), "w": round(s.Width, 1), "h": round(s.Height, 1),
        "z": safe(lambda: s.ZOrderPosition),
        "needs_alt": needs_alt(s),
    }
    rot = safe(lambda: s.Rotation, 0)
    if rot:
        d["rotation"] = round(rot, 1)
    if not safe(lambda: s.Visible, -1):
        d["hidden"] = True
    t = text_of(s)
    if t:
        d["text"] = t[:160].replace("\r", " / ").replace("\x0b", " ")
    ph = placeholder_info(s)
    if ph:
        d["placeholder"] = ph
    if targets.get(s.Id):
        d["animation_effects"] = targets[s.Id]
    if triggers and s.Id in triggers:
        d["animation_trigger"] = True
    if s.Type == 6 and deep:
        d["members"] = [shape_info(s.GroupItems(i), targets, triggers, deep=False)
                        for i in range(1, s.GroupItems.Count + 1)]
        anim_inside = sum(targets.get(m["id"], 0) for m in d["members"])
        if anim_inside:
            d["animated_members"] = anim_inside
    return d


def inspect(pres, slide: int, part: str = "slide", inherited: bool = True) -> dict:
    container = get_part(pres, slide, part)
    targets, triggers = timeline_targets(container)
    out = {
        "where": part_label(pres, slide, part),
        "shapes": [shape_info(container.Shapes(i), targets, triggers) for i in range(1, container.Shapes.Count + 1)],
    }
    if part == "slide":
        sl = container
        out["layout"] = sl.CustomLayout.Name
        out["title"] = slide_title(sl)
        out["effects_total"] = sum(targets.values())
        if inherited:
            out["inherited"] = inherited_objects(pres, slide)
    return out


def inherited_objects(pres, slide: int) -> list[dict]:
    """Non-placeholder objects a slide shows from its layout and master (edit them in Slide Master view)."""
    sl = pres.Slides(slide)
    out = []
    if not safe(lambda: sl.DisplayMasterShapes, -1):  # "Hide background graphics" is ticked
        return out
    lay = sl.CustomLayout
    for s in (lay.Shapes(i) for i in range(1, lay.Shapes.Count + 1)):
        if s.Type != 14:
            out.append({"part": "layout", **shape_info(s)})
    if safe(lambda: lay.DisplayMasterShapes, -1):
        m = sl.Master
        for s in (m.Shapes(i) for i in range(1, m.Shapes.Count + 1)):
            if s.Type != 14:
                out.append({"part": "master", **shape_info(s)})
    return out


def slide_title(sl) -> str | None:
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        if s.Type == 14 and safe(lambda: s.PlaceholderFormat.Type) in TITLE_PLACEHOLDERS:
            return text_of(s).strip()
    return None
