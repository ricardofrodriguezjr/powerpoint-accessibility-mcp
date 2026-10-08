"""Structural fixes that make WCAG grouping possible without changing how the deck looks or animates.

  flatten_group         an existing group that is animated as a whole -> its pieces, each with the same
                        Appear/Fade effect starting together (so it can be nested in a bigger group)
  placeholder_to_picture  a picture sitting in a content placeholder -> an ordinary picture (PowerPoint
                        cannot group placeholders)
  move_to_slides        objects on a slide layout / master -> the slides that show them (screen readers
                        and the Accessibility Checker skip layout/master content)

Every function is a transaction like ops.group_shapes: checkpoint, change, verify (rendered pixels,
positions, animation timing), then save, or reopen the last save on any difference.
"""
from __future__ import annotations

import copy
import time

from . import animations as anim
from . import render
from .com import TRUE, ToolError, com_err, safe
from .inventory import (find_shape, get_part, part_label, slides_showing_master, slides_using_layout,
                        timeline_targets, top_level_index)
from .ops import GEOM_TOLERANCE, VISUAL_TOLERANCE_PX, _geometry, _geom_diffs, _timing_now, _xml_compare, is_junk
from .session import checkpoint, out_dir, require_editable, revert, save, save_copy, stamp, tmp_dir

# Effects for which "animate the group" looks exactly like "animate every piece at the same moment".
FLATTEN_SAFE_EFFECTS = {1: "Appear", 10: "Fade"}
PICTURE_TYPES = {11, 13, 28, 29}  # linked picture, picture, graphic (SVG/icon), linked graphic


def _paste(container, tries: int = 20):
    """Shapes.Paste with retries: the clipboard is sometimes not ready right after Copy()."""
    last = None
    for _ in range(tries):
        try:
            return container.Shapes.Paste()
        except Exception as exc:
            last = exc
            time.sleep(0.3)
    raise ToolError(f"paste failed: {com_err(last)}")


def _finish(pres, result: dict, failures: list[str], autosave: bool) -> dict:
    if failures:
        revert(pres)
        result.update(ok=False, reverted=True, failures=failures,
                      message="Nothing was kept: the deck was reopened from its last save.")
        return result
    if autosave:
        save(pres)
    result.update(ok=True, saved=autosave)
    return result


# --------------------------------------------------------------------------------------------------
# Flatten an animated group
# --------------------------------------------------------------------------------------------------

def check_flatten(container, group_id: int) -> tuple[list[str], dict]:
    """(blockers, info) for flattening group `group_id` on a slide."""
    idx = top_level_index(container, group_id)
    if idx is None:
        return [f"#{group_id} is not a top-level shape on this slide"], {}
    g = container.Shapes(idx)
    if g.Type != 6:
        return [f"#{group_id} is not a group"], {}
    snap = anim.snapshot(container)
    effects = [e for q in snap for e in q["effects"] if e["shape_id"] == group_id]
    blockers = []
    if any(q["trigger_shape"] == group_id for q in snap):
        blockers.append(f"#{group_id} is the click trigger of other animations; ungrouping would remove the trigger")
    pieces = [g.GroupItems(i) for i in range(1, g.GroupItems.Count + 1)]
    targets, triggers = timeline_targets(container)
    own = [pc.Id for pc in pieces if targets.get(pc.Id) or pc.Id in triggers]
    if own:
        blockers.append(f"pieces {own} of #{group_id} have animations of their own; their visibility depends on the "
                        "group's effect, which flattening cannot express")
    for kind_ in (1, 2):  # ppMouseClick, ppMouseOver
        act = safe(lambda k=kind_: g.ActionSettings(k).Action, 0)
        link = safe(lambda k=kind_: g.ActionSettings(k).Hyperlink.Address, "") or \
            safe(lambda k=kind_: g.ActionSettings(k).Hyperlink.SubAddress, "")
        if act or link:
            blockers.append(f"#{group_id} has a click/hover action or hyperlink that ungrouping would drop")
    if any(e["EffectType"] == 10 for e in effects):  # Fade: pieces fade separately, overlaps show through
        from .ops import _bbox, _overlap
        boxes = [(pc.Id, _bbox([pc])) for pc in pieces]
        if any(_overlap(a[1], b[1]) for i, a in enumerate(boxes) for b in boxes[i + 1:]):
            blockers.append(f"#{group_id} fades and its pieces overlap; faded separately they would show through "
                            "each other during the transition (only Appear is safe here)")
    for e in effects:
        if e["EffectType"] not in FLATTEN_SAFE_EFFECTS or e["Paragraph"]:
            blockers.append(f"#{group_id} uses '{e['DisplayName'] or e['EffectType']}' (type {e['EffectType']}); only "
                            f"{'/'.join(FLATTEN_SAFE_EFFECTS.values())} look identical when applied to each piece")
        if e["Timing"].get("AutoReverse") or (e["Timing"].get("RepeatCount") or 1) != 1:
            blockers.append(f"#{group_id} effect repeats or auto-reverses")
    info = {"effects_on_group": len(effects), "pieces": [pc.Id for pc in pieces],
            "group_alt": g.AlternativeText or "", "group_title": safe(lambda: g.Title, "") or "",
            "group_decorative": bool(safe(lambda: g.Decorative, 0))}
    return blockers, info


def flatten_group(pres, slide: int, group_id: int, allow_alt_loss: bool = False, autosave: bool = True) -> dict:
    """Ungroup an existing group that is animated as a whole and give each piece the group's effect,
    all starting together, so the pieces can join a bigger group. The animation looks identical
    (verified with a timeline model and a before/after render)."""
    require_editable(pres)
    container = get_part(pres, slide, "slide")
    blockers, info = check_flatten(container, group_id)
    if blockers:
        raise ToolError("cannot flatten: " + "; ".join(blockers))
    alt = (info["group_alt"] + " " + info["group_title"]).strip()
    if alt and not is_junk(alt) and not allow_alt_loss:
        raise ToolError(f"#{group_id} has its own alt text/title ('{alt[:60]}'); it is lost on ungrouping. Re-run "
                        "with allow_alt_loss=True if the bigger group's alt text covers it")
    pieces = info["pieces"]
    g0 = container.Shapes(top_level_index(container, group_id))
    has_frame = set()  # pieces with a text frame, even an empty one (rectangles, ovals, arrows...)
    for i in range(1, g0.GroupItems.Count + 1):
        it = g0.GroupItems(i)
        if safe(lambda: it.HasTextFrame):
            has_frame.add(it.Id)
    before = anim.snapshot(container)
    expected = copy.deepcopy(before)
    for q in expected:
        new = []
        for e in q["effects"]:
            if e["shape_id"] != group_id:
                new.append(e)
                continue
            for k, pid in enumerate(pieces):
                c = copy.deepcopy(e)
                c["shape_id"] = pid
                c["Paragraph"] = 0
                if pid in has_frame:
                    # a group animates each piece's fill AND text; on a shape with a text frame that is "As
                    # One Object" (animBg=1). Without it PowerPoint animates only the text (nothing, if the
                    # frame is empty) and the shape itself shows from the start.
                    c["AnimateBackground"] = TRUE
                if k:  # every further piece starts together with the first one
                    c["Timing"]["TriggerType"] = anim.WITH_PREV
                    c["Sound"] = {"Type": 0, "Name": None}  # one sound, not one per piece
                new.append(c)
        q["effects"] = new
    model_before = anim.timeline_model(before, {group_id: pieces})

    checkpoint(pres)
    result = {"where": part_label(pres, slide, "slide"), "flattened_group": group_id, "pieces": pieces,
              "effects_on_group": info["effects_on_group"]}
    if info["group_decorative"]:
        result["note_decorative"] = "the group was marked decorative; group the pieces next"
    failures: list[str] = []
    try:
        od = out_dir(pres)
        tag = f"s{slide}_flat{group_id}_{stamp()}"
        img_before = render.export(pres, slide, od / f"{tag}_before.png")
        geom_before = _geometry(container, set(pieces))
        g = container.Shapes(top_level_index(container, group_id))
        g.Ungroup()
        rep = anim.rebuild(container, expected, lambda sid: find_shape(container, sid))
        failures += [f"animation property not restored: {f}" for f in rep["property_failures"]]
        after = anim.snapshot(container)
        failures += [f"animation differs from plan: {d}" for d in anim.compare(expected, after)]
        for q in after:
            for e in q["effects"]:
                if e["shape_id"] in has_frame and e["shape_id"] in pieces and not e.get("AnimateBackground"):
                    failures.append(f"#{e['shape_id']} would animate only its text (shape visible too early)")
        model_after = anim.timeline_model(after)
        if model_after != model_before:
            bad = sorted(k for k in set(model_before) | set(model_after) if model_before.get(k) != model_after.get(k))
            failures.append(f"what the viewer sees would change for shapes {bad}")
        failures += _geom_diffs(geom_before, _geometry(container, set(pieces)))
        vis = render.diff(img_before, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"slide picture changed ({vis['changed_pixels']} px)")
        result.update(rebuild=rep, visual_check=vis, timeline_check="identical" if model_after == model_before else "DIFFERENT",
                      note="group structure removed; pieces keep the same effect, starting together. "
                           "Group them into the larger figure next.")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


# --------------------------------------------------------------------------------------------------
# Placeholder picture -> ordinary picture
# --------------------------------------------------------------------------------------------------

def placeholder_to_picture(pres, slide: int, shape_id: int, autosave: bool = True) -> dict:
    """Replace a picture that sits in a content/picture placeholder with an identical ordinary picture
    (same position, size, crop, stacking order, alt text, animations), so it can be grouped. The old
    placeholder is removed; the new shape gets a new id (returned)."""
    require_editable(pres)
    container = get_part(pres, slide, "slide")
    idx = top_level_index(container, shape_id)
    if idx is None:
        raise ToolError(f"#{shape_id} is not a top-level shape on slide {slide}")
    s = container.Shapes(idx)
    if s.Type != 14:
        raise ToolError(f"#{shape_id} is not a placeholder (type {s.Type}); it can be grouped as it is")
    contained = safe(lambda: s.PlaceholderFormat.ContainedType)
    if contained not in PICTURE_TYPES:
        raise ToolError(f"#{shape_id} holds type {contained}, not a picture; only pictures are converted")
    sl0 = pres.Slides(slide)
    for where, part_obj in (("layout", sl0.CustomLayout), ("master", sl0.Master)):
        tl = safe(lambda part_obj=part_obj: part_obj.TimeLine)
        n = 0
        if tl is not None:
            n = (safe(lambda: tl.MainSequence.Count, 0) or 0) + (safe(lambda: tl.InteractiveSequences.Count, 0) or 0)
        if n:
            raise ToolError(f"the slide's {where} has its own animations, which a placeholder can inherit; "
                            "converting it could drop that animation. Use the carrier technique instead")
    before = anim.snapshot(container)
    slide_id = pres.Slides(slide).SlideID
    props = {"Left": s.Left, "Top": s.Top, "Width": s.Width, "Height": s.Height, "Rotation": safe(lambda: s.Rotation, 0)}
    z0 = s.ZOrderPosition
    keep = {"Name": s.Name, "AlternativeText": s.AlternativeText, "Title": safe(lambda: s.Title, ""),
            "Decorative": safe(lambda: s.Decorative)}

    checkpoint(pres)
    result = {"where": f"slide {slide}", "old_id": shape_id}
    failures: list[str] = []
    try:
        od = out_dir(pres)
        tag = f"s{slide}_ph{shape_id}_{stamp()}"
        img_before = render.export(pres, slide, od / f"{tag}_before.png")
        before_copy = tmp_dir(pres) / f"b{slide_id}.pptx"
        save_copy(pres, str(before_copy))
        s.Copy()
        new = _paste(container)(1)
        if new.Type == 14:
            raise ToolError("PowerPoint pasted it as a placeholder again; cannot convert this one")
        for k, v in props.items():
            setattr(new, k, v)
        while new.ZOrderPosition > z0 + 1:  # just above the old one; it takes its place once that is deleted
            new.ZOrder(3)  # msoSendBackward
        tl = container.TimeLine
        for q in [tl.MainSequence] + [tl.InteractiveSequences(k) for k in range(1, tl.InteractiveSequences.Count + 1)]:
            for i in range(1, q.Count + 1):
                e = q.Item(i)
                if safe(lambda: e.Shape.Id) == shape_id:
                    e.Shape = new
                if safe(lambda: e.Timing.TriggerShape.Id) == shape_id:
                    e.Timing.TriggerShape = new
        new.Name = keep["Name"]
        new.AlternativeText = keep["AlternativeText"]
        if keep["Title"]:
            new.Title = keep["Title"]
        if keep["Decorative"]:
            new.Decorative = TRUE
        new_id = new.Id
        container.Shapes(top_level_index(container, shape_id)).Delete()
        new = find_shape(container, new_id)
        if new.ZOrderPosition != z0:
            failures.append(f"stacking order changed ({z0} -> {new.ZOrderPosition})")
        geo = [k for k, v in props.items() if abs(getattr(new, k) - v) > GEOM_TOLERANCE]
        if geo:
            failures.append(f"geometry changed: {geo}")
        after = anim.snapshot(container)
        remapped = copy.deepcopy(before)
        for q in remapped:
            if q["trigger_shape"] == shape_id:
                q["trigger_shape"] = new_id
            for e in q["effects"]:
                if e["shape_id"] == shape_id:
                    e["shape_id"] = new_id
                if e["Timing"].get("TriggerShape") == shape_id:
                    e["Timing"]["TriggerShape"] = new_id
        com_diffs = anim.compare(remapped, after)
        try:
            xml = _xml_compare(anim.timing_from_pptx(str(before_copy), slide_id, {shape_id: new_id}),
                               _timing_now(pres, slide_id))
        finally:
            safe(lambda: before_copy.unlink())
        if xml["status"] in ("DIFFERENT", "BUILD_LIST_ONLY"):
            failures.append("animation XML differs after replacing the placeholder (see xml_check.diff)")
            failures += [f"animation differs: {d}" for d in com_diffs]
        vis = render.diff(img_before, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"slide picture changed ({vis['changed_pixels']} px)")
        result.update(new_id=new_id, new_type=new.Type, xml_check=xml, visual_check=vis,
                      note="picture is no longer a layout placeholder (looks identical); it can now be grouped")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


# --------------------------------------------------------------------------------------------------
# Layout / master objects -> slides
# --------------------------------------------------------------------------------------------------

def move_to_slides(pres, slide: int, ids: list[int], part: str = "layout", autosave: bool = True) -> dict:
    """Move objects from the slide layout (or master) that `slide` uses onto every slide that shows that
    layout/master: same position, size, stacking (behind the slide's own content), alt text and decorative
    flag; then delete them from the layout/master. Screen readers and the checker then see them."""
    require_editable(pres)
    if part not in ("layout", "master"):
        raise ToolError("part must be 'layout' or 'master'")
    src = get_part(pres, slide, part)
    idx = []
    for sid in ids:
        i = top_level_index(src, sid)
        if i is None:
            raise ToolError(f"#{sid} is not a top-level object on {part_label(pres, slide, part)}")
        if src.Shapes(i).Type == 14:
            raise ToolError(f"#{sid} is a placeholder; layout placeholders are prompts, not content")
        idx.append(i)
    idx.sort(key=lambda i: src.Shapes(i).ZOrderPosition)
    targets_anim, trig_anim = timeline_targets(src)
    moved_ids = set()
    for i in idx:
        sh = src.Shapes(i)
        moved_ids.add(sh.Id)
        if sh.Type == 6:
            moved_ids |= {sh.GroupItems(k).Id for k in range(1, sh.GroupItems.Count + 1)}
    if any(targets_anim.get(x) or x in trig_anim for x in moved_ids):
        raise ToolError("some of these layout objects (or their pieces) are animated or are click triggers; "
                        "moving them is not supported")
    if part == "layout":
        shown = slides_using_layout(pres, src, shown_only=True)
        all_users = slides_using_layout(pres, src)
    else:
        shown = slides_showing_master(pres, pres.Slides(slide).Design.Index)
        all_users = shown
    if not shown:
        raise ToolError("no slide shows these objects (every slide using it hides background graphics)")
    morph = {3953, 3954, 3955}  # Morph by object/word/character: layout objects are background there
    for n in shown:
        for m_ in (n, n + 1):
            if m_ <= pres.Slides.Count and safe(lambda m_=m_: pres.Slides(m_).SlideShowTransition.EntryEffect) in morph:
                raise ToolError(f"slide {m_} uses a Morph transition; moving layout objects onto slides changes how "
                                "Morph treats them. Ask the user before moving.")
    src_anim_before = anim.snapshot(src)
    src_info = [{"id": src.Shapes(i).Id, "box": (src.Shapes(i).Left, src.Shapes(i).Top, src.Shapes(i).Width,
                                                src.Shapes(i).Height),
                 "alt": src.Shapes(i).AlternativeText, "dec": safe(lambda i=i: src.Shapes(i).Decorative)} for i in idx]

    checkpoint(pres)
    result = {"from": part_label(pres, slide, part), "moved": ids, "slides": shown,
              "slides_using_but_hiding_graphics": [n for n in all_users if n not in shown]}
    failures: list[str] = []
    try:
        od = out_dir(pres)
        stmp = stamp()
        before_imgs = {n: render.export(pres, n, od / f"s{n}_move_{stmp}_before.png") for n in shown}
        before_xml = {n: _timing_now(pres, pres.Slides(n).SlideID) for n in shown}
        mapping: dict[int, dict[int, int]] = {}
        for n in shown:
            sl = pres.Slides(n)
            src.Shapes.Range(idx).Copy()
            rng = _paste(sl)
            new = sorted((rng(k) for k in range(1, rng.Count + 1)), key=lambda s: s.ZOrderPosition)
            if len(new) != len(src_info):
                raise ToolError(f"pasted {len(new)} objects on slide {n}, expected {len(src_info)}")
            m = {}
            for info, s in zip(src_info, new):
                for k, v in zip(("Left", "Top", "Width", "Height"), info["box"]):
                    if abs(getattr(s, k) - v) > GEOM_TOLERANCE:
                        setattr(s, k, v)
                s.AlternativeText = info["alt"]
                if info["dec"]:
                    s.Decorative = TRUE
                m[info["id"]] = s.Id
            for s in sorted(new, key=lambda s: -s.ZOrderPosition):  # topmost first -> relative order kept
                s.ZOrder(1)  # msoSendToBack: layout objects were always behind the slide's own content
            mapping[n] = m
        for i in sorted(idx, reverse=True):
            src.Shapes(i).Delete()
        if anim.compare(src_anim_before, anim.snapshot(src)):
            failures.append(f"the {part}'s own animations changed when the objects were removed")
        visual = {}
        for n in shown:
            v = render.diff(before_imgs[n], render.export(pres, n, od / f"s{n}_move_{stmp}_after.png"),
                            od / f"s{n}_move_{stmp}_compare.png")
            visual[n] = {"changed_pixels": v["changed_pixels"], "image": v.get("image")}
            if v["changed_pixels"] > VISUAL_TOLERANCE_PX:
                failures.append(f"slide {n} looks different after the move ({v['changed_pixels']} px)")
            x = _xml_compare(before_xml[n], _timing_now(pres, pres.Slides(n).SlideID))
            if x["status"] in ("DIFFERENT", "BUILD_LIST_ONLY"):
                failures.append(f"slide {n} animation XML changed")
        result.update(mapping={str(k): v for k, v in mapping.items()}, visual_check=visual,
                      note=f"objects now live on the slides; removed from the {part}. Slides added later from "
                           f"this {part} will not have them.",
                      reading_order_note="they sit at the back of each slide's stacking order, so screen readers "
                                         "read them FIRST; run fix_reading_order after all edits (it restacks "
                                         "them invisibly where they belong)")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)
