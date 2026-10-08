"""Edits: alt text, decorative, grouping that keeps animations, restoring animations from another deck.

Every edit that can lose information runs as a transaction: the working copy is saved first, the
edit is verified (animations, geometry, rendered pixels, timing XML), and on any failure the deck is
reopened from disk so the slide is exactly as it was.
"""
from __future__ import annotations

import os
import re

from . import animations as anim
from . import render
from .com import TRUE, FALSE, ToolError, app, com_err, retry, safe
from .inventory import (find_shape, get_part, iter_shapes, kind, part_label, shape_info, timeline_targets,
                        top_level_index)
from .session import (checkpoint, find_open, out_dir, require_editable, revert, save, save_copy, stamp,
                      tmp_dir)

VISUAL_TOLERANCE_PX = 20   # pixels (any channel) allowed to differ in the before/after export
TEXT_BUILD_PROPS = ("AnimateBackground", "BuildByLevelEffect", "TextUnitEffect", "AnimateTextInReverse",
                    "AfterEffect", "Paragraph")
GEOM_TOLERANCE = 0.05      # points

# --------------------------------------------------------------------------------------------------
# Alt text quality
# --------------------------------------------------------------------------------------------------

_JUNK = [
    (re.compile(r"^\s*$"), "empty"),
    (re.compile(r"\.(png|jpe?g|gif|bmp|tiff?|svg|emf|wmf|webp|heic)\b", re.I), "looks like a file name"),
    (re.compile(r"https?://|www\.", re.I), "contains a URL"),
    (re.compile(r"description automatically generated|a picture containing|^a screenshot of a computer",
                re.I), "auto-generated description"),
    (re.compile(r"^(picture|image|graphic|group|rectangle|oval|freeform|content placeholder|text ?box|shape)\s*\d*$",
                re.I), "is just the object name"),
]
_PREFIX = re.compile(r"^\s*(an? )?(image|picture|photo|graphic|icon) (of|showing)\b", re.I)


def lint_alt(text: str | None) -> list[str]:
    text = text or ""
    out = [msg for rx, msg in _JUNK if rx.search(text)]
    if text and _PREFIX.search(text):
        out.append("starts with 'image of'/'picture of' (screen readers already announce the object type)")
    if len(text) > 300:
        out.append(f"{len(text)} characters: consider a shorter alt text plus a longer description in the "
                   "slide notes or slide text")
    return out


def is_junk(text: str | None) -> bool:
    return bool(text and text.strip()) and any(rx.search(text) for rx, _ in _JUNK[1:])


# --------------------------------------------------------------------------------------------------
# Alt text / decorative
# --------------------------------------------------------------------------------------------------

def _touch(pres, slide, part, ids):
    require_editable(pres)
    container = get_part(pres, slide, part)
    return container, [find_shape(container, i) for i in ids]


def set_alt_text(pres, slide: int, ids: list[int], text: str, part: str = "slide") -> dict:
    container, shapes = _touch(pres, slide, part, ids)
    res = []
    for s in shapes:
        before = {"alt": s.AlternativeText, "decorative": safe(lambda: bool(s.Decorative))}
        if safe(lambda: s.Decorative):
            s.Decorative = FALSE
        s.AlternativeText = text
        parent = safe(lambda: s.ParentGroup.Id)
        res.append({"id": s.Id, "name": s.Name, "before": before, "alt": s.AlternativeText,
                    **({"warning": f"#{s.Id} is inside group #{parent}; screen readers read the group's alt "
                                   "text, not this one"} if parent else {})})
    return {"where": part_label(pres, slide, part), "changed": res, "lint": lint_alt(text)}


def mark_decorative(pres, slide: int, ids: list[int], part: str = "slide", decorative: bool = True) -> dict:
    container, shapes = _touch(pres, slide, part, ids)
    res = []
    for s in shapes:
        if safe(lambda: s.Decorative) is None:
            raise ToolError("Shape.Decorative is not available in this PowerPoint build (needs Microsoft 365)")
        before = {"alt": s.AlternativeText, "decorative": bool(s.Decorative)}
        s.Decorative = TRUE if decorative else FALSE
        if decorative and s.AlternativeText:
            s.AlternativeText = ""  # decorative objects must not carry text a screen reader could read
        res.append({"id": s.Id, "name": s.Name, "before": before, "decorative": bool(s.Decorative),
                    "alt": s.AlternativeText})
    return {"where": part_label(pres, slide, part), "changed": res}


def clear_alt_text(pres, slide: int, ids: list[int], part: str = "slide", title_too: bool = True) -> dict:
    container, shapes = _touch(pres, slide, part, ids)
    res = []
    for s in shapes:
        before = {"alt": s.AlternativeText, "title": safe(lambda: s.Title)}
        s.AlternativeText = ""
        if title_too and safe(lambda: s.Title):
            s.Title = ""
        res.append({"id": s.Id, "before": before})
    return {"where": part_label(pres, slide, part), "cleared": res}


# --------------------------------------------------------------------------------------------------
# Grouping
# --------------------------------------------------------------------------------------------------

def _leaf_ids(s) -> set[int]:
    """Ids of the shapes COM can still address after grouping. Shape.GroupItems is flat (nested groups
    are not returned), so a selected group contributes its leaf shapes, not itself."""
    if s.Type == 6:
        return {s.GroupItems(i).Id for i in range(1, s.GroupItems.Count + 1)}
    return {s.Id}


def _bbox(shapes):
    return (min(s.Left for s in shapes), min(s.Top for s in shapes),
            max(s.Left + s.Width for s in shapes), max(s.Top + s.Height for s in shapes))


def _overlap(a, b) -> bool:
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


def _triggers_by_seq(container) -> dict[int, int]:
    """trigger shape id -> number of effects it triggers."""
    out = {}
    tl = safe(lambda: container.TimeLine)
    if tl is None:
        return out
    for k in range(1, tl.InteractiveSequences.Count + 1):
        q = tl.InteractiveSequences(k)
        trig = safe(lambda: q.Item(1).Timing.TriggerShape.Id) if q.Count else None
        if trig is not None:
            out[trig] = out.get(trig, 0) + q.Count
    return out


def check_group(pres, slide: int, ids: list[int], part: str = "slide") -> dict:
    """Dry run: can these shapes be grouped, and what has to be rebuilt?"""
    container = get_part(pres, slide, part)
    blockers, warnings = [], []
    if len(set(ids)) != len(ids) or len(ids) < 2:
        blockers.append("give at least two distinct shape ids")
    shapes = []
    targets, triggers = timeline_targets(container)
    for sid in ids:
        idx = top_level_index(container, sid)
        if idx is None:
            try:
                inner = find_shape(container, sid)
                blockers.append(f"#{sid} is inside group #{inner.ParentGroup.Id}; select the group instead")
            except ToolError:
                blockers.append(f"no shape #{sid} on {part_label(pres, slide, part)}")
            continue
        s = container.Shapes(idx)
        shapes.append(s)
        if s.Type == 14:
            pic = safe(lambda: s.PlaceholderFormat.ContainedType) in (11, 13, 28, 29)
            blockers.append(
                f"#{sid} is a placeholder ({kind(s)}); PowerPoint cannot group placeholders. " +
                (f"RECOMMENDED for WCAG: placeholder_to_picture(slide, {sid}) first (it looks identical), then "
                 f"group with '@{sid}'. Alternative: carrier technique (alt text on #{sid}, other pieces decorative)."
                 if pic else "It is a text/content placeholder: leave it out of the group (its text is read).")
            )
        if s.Type == 19 or safe(lambda: s.HasTable) == TRUE:
            blockers.append(f"#{sid} is a table; PowerPoint cannot group tables (tables are read cell by cell)")
        if s.Type in (4, 23):
            blockers.append(f"#{sid} is a comment/ink comment")
        if s.Type == 6 and (targets.get(sid) or sid in triggers):
            types = sorted({e["EffectType"] for q in anim.snapshot(container) for e in q["effects"] if e["shape_id"] == sid})
            flat_ok = sid not in triggers and types and all(t in (1, 10) for t in types)
            blockers.append(
                f"#{sid} is an existing group animated as a whole; nested inside a new group PowerPoint can no longer "
                "animate it. Options: " +
                (f"(a) RECOMMENDED for WCAG: flatten_group(slide, {sid}) first (its Appear/Fade animation stays "
                 "visually identical, verified), then group everything with one alt text; " if flat_ok else
                 f"(a) flattening is not safe here (effect types {types} would look different per piece); ") +
                f"(b) carrier: full alt text on one piece, other pieces decorative (WCAG-valid, keeps structure); "
                f"(c) leave #{sid} out of the new group; (d) group anyway and lose its animation (changes the design).")
    leaves = set()
    for s in shapes:
        leaves |= _leaf_ids(s)

    effects = sum(n for sid, n in targets.items() if sid in leaves)
    trig_in = sorted(leaves & triggers)
    if trig_in:
        warnings.append(f"shapes {trig_in} are click-triggers for other animations; the trigger sequences are "
                        "re-created on the shapes inside the group and verified")
    if shapes and not blockers:
        zs = [s.ZOrderPosition for s in shapes]
        box = _bbox(shapes)
        overlap, reorder = [], []
        for i in range(1, container.Shapes.Count + 1):
            o = container.Shapes(i)
            if o.Id in ids or not (min(zs) < o.ZOrderPosition < max(zs)):
                continue
            (overlap if _overlap(box, _bbox([o])) else reorder).append(o.Id)
        if overlap:
            warnings.append(f"shapes {overlap} sit between the selected shapes in the stacking order and overlap "
                            "them; grouping lifts the lower selected shapes above them. The before/after picture "
                            "check will catch any visible change and undo it.")
        if reorder:
            warnings.append(f"reading order changes: shapes {reorder} are read before the new group instead of "
                            "between its pieces (no visual change)")
    return {
        "where": part_label(pres, slide, part), "ok": not blockers, "blockers": blockers, "warnings": warnings,
        "members": [shape_info(s, targets, triggers, deep=False) for s in shapes],
        "animation_effects_to_rebuild": effects + sum(_triggers_by_seq(container).get(t, 0) for t in trig_in),
        "note": ("PowerPoint deletes these effects when grouping; group_shapes re-adds them to the same shapes "
                 "inside the group and verifies the result") if effects or trig_in else "no animations involved",
    }


def _geometry(container, ids: set[int]) -> dict:
    out = {}
    for s, _, _ in iter_shapes(container.Shapes):
        if s.Id in ids and s.Type != 6:
            out[s.Id] = (s.Left, s.Top, s.Width, s.Height, safe(lambda: s.Rotation, 0))
    return out


def _geom_diffs(a: dict, b: dict) -> list[str]:
    diffs = []
    for sid, g in a.items():
        h = b.get(sid)
        if h is None:
            diffs.append(f"#{sid} disappeared")
        elif any(abs(x - y) > GEOM_TOLERANCE for x, y in zip(g, h)):
            diffs.append(f"#{sid} moved/resized: {tuple(round(v, 2) for v in g)} -> {tuple(round(v, 2) for v in h)}")
    return diffs


def _timing_now(pres, slide_id: int, folder=None) -> tuple[str, str]:
    """The slide's timing XML exactly as PowerPoint would write it now (SaveCopyAs of the in-memory deck).
    `folder` defaults to the deck's checks folder (pass the working copy's for a source deck, so nothing
    is ever written next to an original)."""
    tmp = str((folder or tmp_dir(pres)) / f"t{os.getpid() % 10000}_{slide_id}.pptx")
    try:
        save_copy(pres, tmp)
        return anim.timing_from_pptx(tmp, slide_id)
    finally:
        safe(lambda: os.remove(tmp))


def _xml_compare(before: tuple, after: tuple) -> dict:
    if before[0] == after[0]:
        return {"status": "identical", "bytes": len(before[0])}
    if before[1] == after[1]:
        return {"status": "equivalent", "note": "same animation; only internal ids/ordering of XML differ"}
    if len(before) > 2 and len(after) > 2 and before[2] == after[2]:
        return {"status": "BUILD_LIST_ONLY", "note": "identical timing tree; only the build list (text-build "
                "bookkeeping) differs. Accepted only if Slide Show playback is identical.",
                "diff": anim.xml_diff(before[1], after[1], limit=20)}
    return {"status": "DIFFERENT", "diff": anim.xml_diff(before[1], after[1])}


def _render_slide_for(pres, slide: int, part: str) -> int | None:
    """A slide whose render shows the edited part (a layout/master edit is invisible on a slide that
    hides background graphics)."""
    if part == "slide":
        return slide
    sl = pres.Slides(slide)
    if safe(lambda: sl.DisplayMasterShapes, -1) and (part == "layout" or safe(lambda: sl.CustomLayout.DisplayMasterShapes, -1)):
        return slide
    for n in range(1, pres.Slides.Count + 1):
        o = pres.Slides(n)
        same = (o.CustomLayout.Design.Index, o.CustomLayout.Index) == (sl.CustomLayout.Design.Index, sl.CustomLayout.Index) \
            if part == "layout" else o.Design.Index == sl.Design.Index
        if same and safe(lambda: o.DisplayMasterShapes, -1):
            return n
    return None


def _tag(slide: int, part: str, what: str, ids: list[int]) -> str:
    # Short on purpose: the checks folder can already be ~190 characters and Windows MAX_PATH is 260.
    return f"s{slide}{part[0]}_{what}{len(ids)}_{stamp()}"


def group_shapes(pres, slide: int, ids: list[int], alt_text: str | None = None, decorative: bool = False,
                 name: str | None = None, part: str = "slide", allow_visual_change: bool = False,
                 autosave: bool = True) -> dict:
    """Group shapes like Ctrl+G, keep every animation effect, put alt text (or decorative) on the group,
    verify, and roll back automatically if anything else changed."""
    require_editable(pres)
    if alt_text and decorative:
        raise ToolError("give alt_text or decorative=True, not both")
    chk = check_group(pres, slide, ids, part)
    if not chk["ok"]:
        raise ToolError("cannot group: " + "; ".join(chk["blockers"]))

    checkpoint(pres)  # the file on disk now equals the "before" state; revert() returns here
    result = {"where": part_label(pres, slide, part), "grouped": ids, "warnings": chk["warnings"]}
    failures: list[str] = []
    try:
        container = get_part(pres, slide, part)
        shapes = [container.Shapes(top_level_index(container, i)) for i in ids]
        leaves = set()
        for s in shapes:
            leaves |= _leaf_ids(s)
        before = anim.snapshot(container)
        n_effects = sum(1 for q in before for e in q["effects"] if e["shape_id"] in leaves) + \
            sum(len(q["effects"]) for q in before if q["trigger_shape"] in leaves)
        geom_before = _geometry(container, leaves)
        od = out_dir(pres)
        tag = _tag(slide, part, "group", ids)
        rs = _render_slide_for(pres, slide, part)
        img_before = render.export(pres, rs, od / f"{tag}_before.png") if rs else None
        slide_id = pres.Slides(slide).SlideID
        before_copy = None
        xml_before = None
        build_groups = None
        whole_fixed: list[int] = []
        if part == "slide":
            before_copy = tmp_dir(pres) / f"b{slide_id}.pptx"
            save_copy(pres, str(before_copy))
            xml_before = anim.timing_from_pptx(str(before_copy), slide_id)
            build_groups = anim.build_groups_from_pptx(str(before_copy), slide_id)
            # COM misreports whole-shape effects written by other apps; correct from the saved targets
            whole_fixed = anim.mark_whole_shape_effects(
                before, anim.whole_targets_from_pptx(str(before_copy), slide_id),
                lambda sid: safe(lambda: find_shape(container, sid).HasTextFrame, 0))
        if whole_fixed:
            result["whole_shape_effects"] = (f"{len(whole_fixed)} effects target the whole shape in the saved file "
                                             "(written by another app); re-created 'as one object'")

        idx = [top_level_index(container, i) for i in ids]
        try:
            g = container.Shapes.Range(idx).Group()
        except Exception as exc:
            raise ToolError(f"PowerPoint refused to group {ids}: {com_err(exc)}")
        gid = g.Id
        rebuild = None
        if n_effects:
            rebuild = anim.rebuild(container, before, lambda sid: find_shape(container, sid), build_groups)
            failures += [f"animation property not restored: {f}" for f in rebuild["property_failures"]]
        g = find_shape(container, gid)
        if name:
            g.Name = name
        if decorative:
            g.Decorative = TRUE
            g.AlternativeText = ""
        elif alt_text is not None:
            g.AlternativeText = alt_text

        after = anim.snapshot(container)
        com_diffs = anim.compare(before, after)
        failures += _geom_diffs(geom_before, _geometry(container, leaves))
        vis = None
        if rs:
            img_after = render.export(pres, rs, od / f"{tag}_after.png")
            vis = render.diff(img_before, img_after, od / f"{tag}_compare.png")
            if vis["changed_pixels"] > VISUAL_TOLERANCE_PX and not allow_visual_change:
                failures.append(f"the slide looks different after grouping ({vis['changed_pixels']} pixels "
                                f"changed, see {vis['image']})")
        else:
            result["visual_check_note"] = "no slide displays this layout/master; visual check skipped"
        if xml_before is not None:
            xml = _xml_compare(xml_before, _timing_now(pres, slide_id))
        else:
            xml = {"status": "skipped (layout/master: COM comparison used)"}
        if xml["status"] == "BUILD_LIST_ONLY":
            # Accept only if no text-build property differs (COM reads these reliably) and no trigger
            # sequence is involved (playback cannot fire triggers); then require identical playback.
            sem = [d for d in com_diffs if any(f"): {k} " in d for k in TEXT_BUILD_PROPS)]
            if sem:
                failures += [f"text-build setting changed: {d}" for d in sem]
            if any(q["seq"] != "main" and any(e["shape_id"] in leaves for e in q["effects"]) for q in before):
                failures.append("build list differs on a slide whose members have trigger animations; "
                                "playback cannot verify triggers")
        if xml["status"] == "BUILD_LIST_ONLY" and not failures:
            from . import playback  # screen capture; only needed in this rare case
            pb = playback.compare_playback(pres, slide, str(before_copy))
            xml["playback"] = {"ok": pb["ok"], "clicks": pb["clicks"], "image": pb["image"],
                               "changed_px": [f["changed_px"] for f in pb["frames"]]}
            if pb["ok"]:
                xml["status"] = "equivalent (playback identical)"
            else:
                failures.append(f"only the build list differs, but Slide Show playback differs (see {pb['image']})")
        if xml["status"] == "DIFFERENT":
            failures.append("saved animation XML differs from the original (see xml_check.diff)")
            failures += [f"animation differs: {d}" for d in com_diffs]
        elif xml["status"].startswith("skipped"):
            failures += [f"animation differs: {d}" for d in com_diffs]  # no XML check: COM view is the authority
        elif com_diffs:
            xml["com_read_differences_ignored"] = com_diffs[:10]  # saved XML matches; COM reads are noisy
        result.update({
            "group_id": gid, "group_name": safe(lambda: find_shape(container, gid).Name),
            "alt_text": alt_text, "decorative": decorative, "lint": lint_alt(alt_text) if alt_text else [],
            "animation_effects_on_members": n_effects, "rebuild": rebuild, "xml_check": xml, "visual_check": vis,
        })
        if part == "slide":
            result["reading_order_note"] = (
                f"the group is read where its front-most piece was (stacking position "
                f"{safe(lambda: find_shape(container, gid).ZOrderPosition)}); run fix_reading_order after all edits")
    except Exception as exc:  # anything at all after the checkpoint: undo it
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    finally:
        bc = locals().get("before_copy")
        if bc:
            safe(lambda: os.remove(str(bc)))
    if failures:
        revert(pres)
        result.update(ok=False, reverted=True, failures=failures,
                      message="Nothing was kept: the deck was reopened from its last save. "
                              "Report the failures; do not retry blindly.")
        return result
    if autosave:
        save(pres)
    result.update(ok=True, saved=autosave)
    return result


# --------------------------------------------------------------------------------------------------
# Restore animations that were lost earlier (e.g. someone grouped by hand)
# --------------------------------------------------------------------------------------------------

def _shape_sig(s) -> tuple:
    t = ""
    if safe(lambda: s.HasTextFrame) and safe(lambda: s.TextFrame.HasText):
        t = (safe(lambda: s.TextFrame.TextRange.Text, "") or "")[:60]
    return (s.Type, t)


def restore_animations_from(pres, slide: int, source_deck: str, source_slide: int | None = None,
                            replace: bool = False, restore_positions: bool = False, autosave: bool = True) -> dict:
    """Copy the animation sequence of `source_deck` slide `source_slide` onto the same shape ids in this
    deck's slide (shapes may now be inside groups), verify, and roll back on any mismatch."""
    require_editable(pres)
    source_slide = source_slide or slide
    a = app()
    src = find_open(source_deck)
    opened = src is None
    if opened:
        src = a.Presentations.Open(os.path.abspath(source_deck), TRUE, FALSE, FALSE)
    od = out_dir(pres)
    tag = f"s{slide}_restore_{stamp()}"
    try:
        src_slide = src.Slides(source_slide)
        src_snap = anim.snapshot(src_slide)
        src_geom = {s.Id: (_shape_sig(s), (s.Left, s.Top, s.Width, s.Height))
                    for s, _, _ in iter_shapes(src_slide.Shapes) if s.Type != 6}
        src_ids = {e["shape_id"] for q in src_snap for e in q["effects"]} | \
                  {q["trigger_shape"] for q in src_snap if q["trigger_shape"]}
        src_timing = _timing_now(src, src_slide.SlideID, tmp_dir(pres))  # PowerPoint's own serialization
        wtmp = str(tmp_dir(pres) / f"w{os.getpid() % 10000}_{src_slide.SlideID}.pptx")
        try:
            save_copy(src, wtmp)
            src_whole = anim.whole_targets_from_pptx(wtmp, src_slide.SlideID)
        finally:
            safe(lambda: os.remove(wtmp))
        img_src = render.export(src, source_slide, od / f"{tag}_source.png")
    finally:
        if opened:
            retry(src.Close)
    if not any(q["effects"] for q in src_snap):
        raise ToolError(f"source slide {source_slide} has no animations")

    container = get_part(pres, slide, "slide")
    problems, moved = [], {}
    for sid in src_ids:
        try:
            s = find_shape(container, sid)
        except ToolError:
            problems.append(f"#{sid} is missing on target slide {slide}")
            continue
        if sid in src_geom and _shape_sig(s) != src_geom[sid][0]:
            problems.append(f"#{sid} differs: source {src_geom[sid][0]} vs target {_shape_sig(s)}")
    if problems:
        raise ToolError("source and target slides do not match: " + "; ".join(problems))
    # COM misreports whole-shape effects written by other apps; correct from the source's saved targets
    anim.mark_whole_shape_effects(src_snap, src_whole,
                                  lambda sid: safe(lambda: find_shape(container, sid).HasTextFrame, 0))
    for s, _, _ in iter_shapes(container.Shapes):  # every leaf shape that also exists in the source
        if s.Type != 6 and s.Id in src_geom:
            box = src_geom[s.Id][1]
            d = (s.Left - box[0], s.Top - box[1])  # exact: a rounded shift leaves a visible sub-point offset
            if abs(d[0]) > GEOM_TOLERANCE or abs(d[1]) > GEOM_TOLERANCE:
                moved[s.Id] = d
    existing = anim.snapshot(container)
    n_existing = sum(len(q["effects"]) for q in existing)
    if n_existing and not replace:
        raise ToolError(f"target slide already has {n_existing} effects; pass replace=True to discard them")

    checkpoint(pres)
    result = {"slide": slide, "source": f"{source_deck} slide {source_slide}",
              "shapes_moved_since_source_pt": {k: (round(v[0], 3), round(v[1], 3)) for k, v in moved.items()}}
    failures: list[str] = []
    try:
        img_before = render.export(pres, slide, od / f"{tag}_before.png")
        tl = container.TimeLine
        if n_existing:
            for k in range(tl.InteractiveSequences.Count, 0, -1):
                q = tl.InteractiveSequences(k)
                for i in range(q.Count, 0, -1):
                    q.Item(i).Delete()
            for i in range(tl.MainSequence.Count, 0, -1):
                tl.MainSequence.Item(i).Delete()
        rep = anim.rebuild(container, src_snap, lambda sid: find_shape(container, sid))
        failures += [f"animation property not restored: {f}" for f in rep["property_failures"]]
        after = anim.snapshot(container)
        com_diffs = anim.compare(src_snap, after)

        position_fix = None
        if restore_positions and moved:
            ref = next(iter(moved.values()))
            same = all(abs(d[0] - ref[0]) < 0.01 and abs(d[1] - ref[1]) < 0.01 for d in moved.values())
            deltas = {ref} if same else set(moved.values())
            parents = {safe(lambda sid=sid: find_shape(container, sid).ParentGroup.Id) for sid in moved}
            group_leaves = set()
            if len(parents) == 1 and None not in parents:
                g = find_shape(container, next(iter(parents)))
                group_leaves = _leaf_ids(g)
            if len(deltas) == 1 and group_leaves and set(moved) == {i for i in group_leaves if i in src_geom}:
                dx, dy = deltas.pop()
                g.Left, g.Top = g.Left - dx, g.Top - dy
                position_fix = {"group": g.Id, "moved_back_by_pt": (round(-dx, 4), round(-dy, 4))}
            else:
                failures.append(f"cannot restore positions automatically: shapes moved by different amounts or "
                                f"are not exactly one group's members ({moved})")

        xml = _xml_compare(src_timing, _timing_now(pres, pres.Slides(slide).SlideID))
        if xml["status"] == "BUILD_LIST_ONLY":
            sem = [d for d in com_diffs if any(f"): {k} " in d for k in TEXT_BUILD_PROPS)]
            failures += [f"text-build setting differs from source: {d}" for d in sem]
        if xml["status"] == "BUILD_LIST_ONLY" and not sem:
            from . import playback
            pb = playback.compare_playback(pres, slide, source_deck, source_slide)
            xml["playback"] = {"ok": pb["ok"], "image": pb["image"]}
            if pb["ok"]:
                xml["status"] = "equivalent (playback identical)"
            else:
                failures.append(f"Slide Show playback differs from the source (see {pb['image']})")
        if xml["status"] == "DIFFERENT":
            failures.append("saved animation XML differs from the source slide")
            failures += [f"animation differs from source: {d}" for d in com_diffs]
        elif com_diffs:
            xml["com_read_differences_ignored"] = com_diffs[:10]
        img_after = render.export(pres, slide, od / f"{tag}_after.png")
        if position_fix:  # moved back: the slide must now look like the source slide
            vis = render.diff(img_src, img_after, od / f"{tag}_compare_source.png")
            vis["compared_with"] = "source slide"
        else:  # nothing moved: the static picture must be unchanged
            vis = render.diff(img_before, img_after, od / f"{tag}_compare.png")
            vis["compared_with"] = "before"
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"slide picture differs from the {vis['compared_with']} ({vis['changed_pixels']} px)")
        result.update({"effects": anim.summarize(after), "rebuild": rep, "xml_check": xml, "visual_check": vis,
                       "position_fix": position_fix})
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    if failures:
        revert(pres)
        result.update(ok=False, reverted=True, failures=failures)
        return result
    if autosave:
        save(pres)
    result.update(ok=True, saved=autosave)
    return result
