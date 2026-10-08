"""Slide titles. Every slide needs a real title placeholder with text (screen readers announce it when a
slide opens and list slides by it; PowerPoint's checker: Error), and titles should be unique
(Section 508 / HHS: "Topic (n of N)"; PowerPoint's checker: "Duplicate slide title").

  promote_to_title   the visible heading (a text box or body placeholder) becomes the slide's title
                     placeholder, with identical position and formatting (looks the same; verified)
  add_hidden_title   fallback: a title placeholder off the slide (Microsoft's "Add Hidden Slide Title")
  set_slide_title    change the title text (e.g. add "(2 of 3)"), keeping its formatting
  make_titles_unique "(n of N)" on every run of slides that share a title
"""
from __future__ import annotations

import re

from . import render
from .com import TITLE_PLACEHOLDERS, ToolError, com_err, safe
from .inventory import get_part, slides_using_layout, text_of, timeline_targets, top_level_index
from .session import checkpoint, out_dir, require_editable, revert, save, stamp
from .textfmt import actions as _actions
from .textfmt import copy_actions, copy_shape_format, copy_text


def title_shape(sl):
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        if s.Type == 14 and safe(lambda: s.PlaceholderFormat.Type) in TITLE_PLACEHOLDERS:
            return s
    return None


def title_text(sl) -> str:
    """The slide title as screen readers get it: empty when there is no title placeholder, or it is empty,
    marked decorative or hidden."""
    t = title_shape(sl)
    if t is None or safe(lambda: t.Decorative, 0) == -1 or safe(lambda: t.Visible, -1) == 0:
        return ""
    return text_of(t).strip()


def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t.replace("\x0b", " ").replace("\r", " ")).strip().casefold()


def _layout_title(lay):
    for i in range(1, lay.Shapes.Count + 1):
        s = lay.Shapes(i)
        if s.Type == 14 and safe(lambda: s.PlaceholderFormat.Type) in TITLE_PLACEHOLDERS:
            return s
    return None


def heading_candidates(sl) -> list[dict]:
    """Text objects that look like the slide's visible heading, best first: short, large, near the top."""
    out = []
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        if s.Type == 14 and safe(lambda: s.PlaceholderFormat.Type) in {13, 14, 15, 16} | set(TITLE_PLACEHOLDERS):
            continue
        if safe(lambda: s.Visible, -1) == 0 or safe(lambda: s.Decorative, 0) == -1:
            continue  # hidden or decorative text is not a visible heading
        if s.Type not in (1, 14, 17) or not (safe(lambda: s.HasTextFrame) and safe(lambda: s.TextFrame.HasText)):
            continue
        txt = text_of(s).strip()
        words = len(txt.split())
        if not txt or words > 20 or "\r" in txt.strip("\r") and words > 12:
            continue
        size = safe(lambda: s.TextFrame2.TextRange.Font.Size, 0) or 0
        if size <= 0 or size > 400:
            size = 18
        bold = safe(lambda: s.TextFrame2.TextRange.Font.Bold, 0) == -1
        score = size * 2 + (6 if bold else 0) - s.Top / 30 - words * 0.5
        out.append({"id": s.Id, "name": s.Name, "text": txt[:120], "font_size": size, "bold": bold,
                    "top": round(s.Top, 1), "score": round(score, 1)})
    return sorted(out, key=lambda d: -d["score"])


def title_report(pres) -> dict:
    missing, groups = [], {}
    for n in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(n)
        t = title_text(sl)
        if not t:
            c = heading_candidates(sl)
            ts = title_shape(sl)
            state = ("absent" if ts is None else "decorative" if safe(lambda: ts.Decorative, 0) == -1 else
                     "hidden" if safe(lambda: ts.Visible, -1) == 0 else "empty")
            nxt = {"decorative": f"mark_decorative(slide, [{ts.Id if ts else 0}], decorative=False): the title is "
                                 "there but marked decorative",
                   "hidden": "un-hide the title placeholder (Selection Pane) or add_hidden_title",
                   }.get(state, "promote_to_title(slide, candidate.id)" if c else "add_hidden_title(slide, text)")
            missing.append({"slide": n, "title_state": state, "next_step": nxt,
                            "layout": sl.CustomLayout.Name,
                            "layout_has_title": _layout_title(sl.CustomLayout) is not None,
                            "candidate": c[0] if c else None, "other_candidates": c[1:3]})
        else:
            groups.setdefault(_norm(t), []).append(n)
    dups = [{"title": title_text(pres.Slides(v[0])), "slides": v} for v in groups.values() if len(v) > 1]
    return {"missing": missing, "duplicates": dups,
            "summary": {"slides_without_title": [m["slide"] for m in missing],
                        "duplicate_title_groups": [d["slides"] for d in dups]}}


def _finish(pres, result, failures, autosave):
    if failures:
        revert(pres)
        result.update(ok=False, reverted=True, failures=failures,
                      message="Nothing was kept: the deck was reopened from its last save.")
        return result
    if autosave:
        save(pres)
    result.update(ok=True, saved=autosave)
    return result


def _ensure_layout_title(pres, sl, result: dict) -> list[tuple]:
    """Give the slide's layout a title placeholder if it has none (Slide Master > Title checkbox). Returns
    before-renders of the other slides using that layout, to prove they look the same afterwards."""
    lay = sl.CustomLayout
    if _layout_title(lay) is not None:
        return []
    od, tag = out_dir(pres), stamp()
    others = [n for n in slides_using_layout(pres, lay) if n != sl.SlideIndex]
    before = [(n, render.export(pres, n, od / f"s{n}_laytitle_{tag}_before.png")) for n in others]
    try:
        lay.Shapes.AddTitle()
    except Exception as exc:
        raise ToolError(f"cannot add a title placeholder to layout '{lay.Name}': {com_err(exc)}")
    result["layout_change"] = (f"layout '{lay.Name}' got a title placeholder (as with Slide Master > Title). Layout "
                               "placeholders are prompts only: slides that don't use it look the same.")
    return before


def _check_others(pres, before: list[tuple], failures: list) -> dict:
    from .ops import VISUAL_TOLERANCE_PX
    out = {}
    od, tag = out_dir(pres), stamp()
    for n, img in before:
        v = render.diff(img, render.export(pres, n, od / f"s{n}_laytitle_{tag}_after.png"))
        out[n] = v["changed_pixels"]
        if v["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"slide {n} (same layout) looks different after the layout change ({v['changed_pixels']} px)")
    return out


def _get_or_add_title(sl, log: list):
    """The slide's title placeholder, adding it from the layout when missing. Shapes.AddTitle is refused on
    slides whose layout type is Blank; re-applying the (now titled) layout adds the empty title placeholder
    instead (empty placeholders do not render)."""
    t = title_shape(sl)
    if t is not None:
        return t
    try:
        return sl.Shapes.AddTitle()
    except Exception as exc:  # noqa: BLE001
        log.append(f"AddTitle refused ({com_err(exc)[:60]}); re-applied layout '{sl.CustomLayout.Name}' instead")
    sl.CustomLayout = sl.CustomLayout
    t = title_shape(sl)
    if t is None:
        raise ToolError("PowerPoint would not add a title placeholder to this slide")
    return t


def _remove_leftover_placeholder(sl, ids_before: set[int], box: tuple, log: list) -> None:
    """Deleting a placeholder that holds text leaves an empty copy of it (like pressing Delete once in the
    UI). Remove that copy too: an empty placeholder with a fill would still paint."""
    for i in range(sl.Shapes.Count, 0, -1):
        o = sl.Shapes(i)
        if o.Id in ids_before or o.Type != 14 or safe(lambda: o.TextFrame.HasText):
            continue
        if abs(o.Left - box[0]) < 1 and abs(o.Top - box[1]) < 1:
            log.append(f"removed the empty placeholder #{o.Id} PowerPoint left behind")
            o.Delete()


def promote_to_title(pres, slide: int, shape_id: int, autosave: bool = True) -> dict:
    """Make the visible heading `shape_id` (text box, text placeholder or text shape) the slide's real title:
    a title placeholder with the same text, position and formatting replaces it. Looks identical (verified
    by render); the layout gets a title placeholder first if it has none. Rolls back on any difference."""
    from .ops import VISUAL_TOLERANCE_PX
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    idx = top_level_index(sl, shape_id)
    if idx is None:
        raise ToolError(f"#{shape_id} is not a top-level shape on slide {slide}")
    src = sl.Shapes(idx)
    if title_text(sl):
        raise ToolError(f"slide {slide} already has a title ('{title_text(sl)[:60]}')")
    if src.Type == 14 and safe(lambda: src.PlaceholderFormat.Type) in TITLE_PLACEHOLDERS:
        raise ToolError(f"#{shape_id} is already the title placeholder")
    if not (safe(lambda: src.HasTextFrame) and safe(lambda: src.TextFrame.HasText)):
        raise ToolError(f"#{shape_id} has no text")
    targets, triggers = timeline_targets(sl)
    if targets.get(shape_id) or shape_id in triggers:
        raise ToolError(f"#{shape_id} is animated; use add_hidden_title(slide, text) instead (keeps its animation)")
    existing = title_shape(sl)
    if existing is not None and (targets.get(existing.Id) or existing.Id in triggers):
        raise ToolError(f"the slide's empty title placeholder #{existing.Id} is animated; reusing it would animate "
                        "the heading. Remove that effect first or use add_hidden_title")
    text = text_of(src)
    actions_before = _actions(src)
    checkpoint(pres)
    result = {"slide": slide, "from": f"#{shape_id} {src.Name}", "text": text.strip()[:120]}
    failures, log = [], []
    try:
        from .readingorder import _restack
        od, tag = out_dir(pres), f"s{slide}_title_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        others = _ensure_layout_title(pres, sl, result)
        t = _get_or_add_title(sl, log)
        safe(lambda: setattr(t, "Decorative", 0))
        safe(lambda: setattr(t, "Visible", -1))
        src = sl.Shapes(top_level_index(sl, shape_id))
        ast = safe(lambda: src.AutoShapeType, 1)
        if ast not in (1, -2, None):  # heading drawn in a rounded/chevron/... shape: same geometry
            t.AutoShapeType = ast
            for i in range(1, (safe(lambda: src.Adjustments.Count, 0) or 0) + 1):
                safe(lambda i=i: t.Adjustments.SetItem(i, src.Adjustments.Item(i)))
        for k in ("HorizontalFlip", "VerticalFlip"):
            if safe(lambda k=k: getattr(src, k)) == -1 and safe(lambda k=k: getattr(t, k)) != -1:
                t.Flip(0 if k == "HorizontalFlip" else 1)
        for k in ("Left", "Top", "Width", "Height"):
            setattr(t, k, getattr(src, k))
        safe(lambda: setattr(t, "Rotation", src.Rotation))
        copy_shape_format(src, t)
        copy_text(src, t, log)
        copy_actions(src, t, log)
        for k in ("Left", "Top", "Width", "Height"):  # Apply() can resize; geometry wins
            setattr(t, k, getattr(src, k))
        if (src.AlternativeText or "").strip():
            t.AlternativeText = src.AlternativeText
        # stack the title exactly where the heading is (just above it, in either direction)
        cur = [sl.Shapes(i).Id for i in range(1, sl.Shapes.Count + 1)]
        new_id = t.Id
        rest = [x for x in cur if x != new_id]
        k = rest.index(shape_id)
        _restack(sl, rest[:k + 1] + [new_id] + rest[k + 1:])
        box = (src.Left, src.Top)
        ids_before = {sl.Shapes(i).Id for i in range(1, sl.Shapes.Count + 1)}
        was_placeholder = src.Type == 14
        src.Delete()
        if was_placeholder:
            _remove_leftover_placeholder(sl, ids_before, box, log)
        vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"the slide looks different with the title placeholder ({vis['changed_pixels']} px, see "
                            f"{vis.get('image')}). Use add_hidden_title instead.")
        if _norm(title_text(sl)) != _norm(text):
            failures.append("title text differs from the heading")
        lost = [a for a in actions_before if a not in _actions(title_shape(sl))]
        if lost:
            failures.append(f"click/hover actions or links of the heading were not carried over: {lost}")
        if others:
            result["other_slides_on_layout_changed_px"] = _check_others(pres, others, failures)
        result.update(title_id=new_id, visual_check=vis, refused_settings=log or None,
                      note="the heading is now the slide title (same look); run fix_reading_order so it is read first")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


def add_hidden_title(pres, slide: int, text: str, autosave: bool = True) -> dict:
    """Microsoft's "Add Hidden Slide Title": a title placeholder holding `text`, placed just above the slide
    so it is not shown in Slide Show (screen readers and the slide list still get the title). Use only when
    no visible heading can be promoted. Verified: the slide renders identically."""
    from .ops import VISUAL_TOLERANCE_PX
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    if title_text(sl):
        raise ToolError(f"slide {slide} already has a title")
    if not text.strip():
        raise ToolError("give the title text")
    checkpoint(pres)
    result = {"slide": slide, "text": text}
    failures = []
    try:
        od, tag = out_dir(pres), f"s{slide}_htitle_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        others = _ensure_layout_title(pres, sl, result)
        log: list = []
        t = _get_or_add_title(sl, log)
        if log:
            result["log"] = log
        tg, tt = timeline_targets(sl)
        if tg.get(t.Id) or t.Id in tt:
            raise ToolError(f"the slide's empty title placeholder #{t.Id} is animated; remove that effect first")
        safe(lambda: setattr(t, "Decorative", 0))
        safe(lambda: setattr(t, "Visible", -1))
        t.TextFrame.TextRange.Text = text
        t.Left = 0
        t.Top = -t.Height - 24  # above the slide: off-slide titles are not shown in Slide Show
        t.Name = "Hidden title"
        vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"slide looks different ({vis['changed_pixels']} px)")
        if others:
            result["other_slides_on_layout_changed_px"] = _check_others(pres, others, failures)
        result.update(title_id=t.Id, visual_check=vis,
                      note="off-slide title (not shown in Slide Show). If the same words are visible on the slide, "
                           "screen readers read them twice; prefer promote_to_title when there is a visible heading.")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


def _u16(s: str) -> int:
    """Length in UTF-16 units: PowerPoint's character positions count those (an astral character such as a
    math-italic letter is 2), Python's str counts code points."""
    return len(s.encode("utf-16-le")) // 2


def _overlaps(a, b) -> bool:
    return min(a.Left + a.Width, b.Left + b.Width) - max(a.Left, b.Left) > 1 and \
        min(a.Top + a.Height, b.Top + b.Height) - max(a.Top, b.Top) > 1


def _title_problem(sl, slide: int) -> str | None:
    """Why set_slide_title cannot make this slide's title what screen readers announce (None = it can)."""
    t = title_shape(sl)
    if t is None:
        return f"slide {slide} has no title placeholder (promote_to_title / add_hidden_title first)"
    if safe(lambda: t.Decorative, 0) == -1:
        return (f"the title placeholder #{t.Id} on slide {slide} is marked decorative, so screen readers skip it: "
                f"mark_decorative(slide={slide}, ids=[{t.Id}], decorative=False) first")
    if safe(lambda: t.Visible, -1) == 0:
        return (f"the title placeholder #{t.Id} on slide {slide} is hidden (Selection Pane): un-hide it, or use "
                "add_hidden_title")
    if not text_of(t).strip():
        over = [c for c in heading_candidates(sl) if (i := top_level_index(sl, c["id"])) and _overlaps(sl.Shapes(i), t)]
        if over:
            return (f"the empty title placeholder #{t.Id} lies over the visible heading #{over[0]['id']}; writing into "
                    f"it would draw a second heading: promote_to_title(slide={slide}, shape_id={over[0]['id']})")
        return (f"the title placeholder #{t.Id} on slide {slide} is empty; text written into it would appear on the "
                "slide (a visible change): use promote_to_title for a visible heading, or add_hidden_title")
    return None


def _retitle(pres, slide: int, text: str, result: dict, failures: list) -> None:
    """The title edit itself, verified, with no checkpoint and no save (callers own the transaction)."""
    sl = pres.Slides(slide)
    t = title_shape(sl)
    full = text_of(t)
    base = full.strip()
    result.update(before=base, after=text)
    od, tag = out_dir(pres), f"s{slide}_settitle_{stamp()}"
    img0 = render.export(pres, slide, od / f"{tag}_before.png")
    box0 = (t.Left, t.Top, t.Width, t.Height)
    tr = t.TextFrame.TextRange
    paras0 = tr.Paragraphs().Count
    warnings = []
    if base and text.startswith(base) and len(text) > len(base):
        off = full.find(base)
        added = tr.Characters(_u16(full[:off]) + 1, _u16(base)).InsertAfter(text[len(base):])
        safe(lambda: setattr(added.Font, "BaselineOffset", 0))
        if safe(lambda: added.ActionSettings(1).Action) == 7:
            safe(lambda: added.ActionSettings(1).Hyperlink.Delete())
        if safe(lambda: added.Font.BaselineOffset, 0):
            failures.append(f"slide {slide}: the added words came out raised/lowered")
        if tr.Paragraphs().Count != paras0:
            failures.append(f"slide {slide}: the added words started a new paragraph")
    else:
        if safe(lambda: tr.Runs().Count, 1) > 1:
            result["formatting_reset"] = "the title had mixed formatting; the new text uses its first run's"
        tr.Text = text
    if _norm(title_text(sl)) != _norm(text):
        failures.append(f"slide {slide}: the title screen readers get is '{title_text(sl)}', not '{text}'")
    vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
    scale = render.WIDTH / pres.PageSetup.SlideWidth
    l0, t0, w0, h0 = box0
    box = vis.get("changed_box_px")
    if box and not (box[0] >= l0 * scale - 4 and box[1] >= t0 * scale - 4 and box[2] <= (l0 + w0) * scale + 4
                    and box[3] <= (t0 + h0) * scale + 4):
        warnings.append("pixels outside the original title box changed (the longer title wraps or overflows); "
                        "check the compare image")
    if t.Height > h0 + 1 or t.Width > w0 + 1:
        warnings.append(f"the title box grew from {w0:.0f}x{h0:.0f} to {t.Width:.0f}x{t.Height:.0f} pt (shape to fit); "
                        "check it does not cover slide content")
    if safe(lambda: t.TextFrame2.AutoSize, 0) == 0 and (safe(lambda: tr.BoundHeight, 0) or 0) > t.Height + 1:
        warnings.append("the new title is taller than its box (overflows); consider a shorter suffix")
    result["visual_check"] = vis
    if warnings:
        result["warnings"] = warnings


def set_slide_title(pres, slide: int, text: str, autosave: bool = True) -> dict:
    """Change the title text keeping its formatting: when the new title extends the current one, only the
    extra words are inserted right after the last visible character (plain, not raised/lowered, not linked);
    otherwise the text is replaced (reported when that resets mixed formatting). Visible change, limited to
    the title. Refuses titles screen readers would still not get (hidden, decorative) and empty placeholders
    (writing into those adds visible text: promote_to_title / add_hidden_title)."""
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    if not text.strip():
        raise ToolError("give the new title text")
    why = _title_problem(sl, slide)
    if why:
        raise ToolError(why)
    checkpoint(pres)
    result = {"slide": slide}
    failures = []
    try:
        _retitle(pres, slide, text, result, failures)
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


def _runs(nums: list[int]) -> list[list[int]]:
    runs, cur = [], [nums[0]]
    for n in nums[1:]:
        if n == cur[-1] + 1:
            cur.append(n)
        else:
            runs.append(cur)
            cur = [n]
    runs.append(cur)
    return runs


def make_titles_unique(pres, slides: list[int] | None = None, pattern: str = "{title} ({n} of {N})") -> dict:
    """Each run of CONSECUTIVE slides sharing a title (a topic continued over several slides) gets
    "(1 of N)" ... "(N of N)" (Section 508 / HHS pattern), or `pattern` with {title}, {n}, {N}. Slides that
    share a title but are not next to each other are not numbered (that would claim a series): they are
    returned in `needs_distinct_title` for a descriptive title (set_slide_title). A run is numbered whole or
    not at all: when `slides` names any slide of a run, the whole run is done, as one verified transaction
    (one failure reverts the run). Every new title is checked against all titles first."""
    from . import session
    require_editable(pres)
    path = pres.FullName
    rep = title_report(pres)
    existing = {_norm(title_text(pres.Slides(n))) for n in range(1, pres.Slides.Count + 1)}
    want = {int(n) for n in slides} if slides else None
    plans, out = [], {"changed": [], "failed": [], "needs_distinct_title": []}
    for d in rep["duplicates"]:
        for run in _runs(d["slides"]):
            if len(run) < 2:
                if want is None or run[0] in want:
                    out["needs_distinct_title"].append({"slide": run[0], "title": d["title"],
                                                        "same_title_as": [n for n in d["slides"] if n != run[0]]})
                continue
            if want is not None and not want & set(run):
                continue
            new = [(n, pattern.format(title=title_text(pres.Slides(n)), n=k, N=len(run))) for k, n in enumerate(run, 1)]
            clash = [t for _, t in new if _norm(t) in existing]
            if len({_norm(t) for _, t in new}) < len(new) or clash:
                out["failed"].append({"slides": run, "titles": [t for _, t in new],
                                      "failures": [f"title(s) already in the deck: {clash}" if clash else
                                                   "the pattern gives the same title twice"]})
                continue
            bad = [why for n, _ in new if (why := _title_problem(pres.Slides(n), n))]
            if bad:
                out["failed"].append({"slides": run, "failures": bad})
                continue
            existing |= {_norm(t) for _, t in new}
            plans.append((run, new))
            if want is not None and not set(run) <= want:
                out.setdefault("extended_to_whole_run", []).append(run)
    for run, new in plans:
        p = session.find_open(path)
        session.checkpoint(p)
        results, failures = [], []
        try:
            for n, title in new:
                r = {"slide": n}
                _retitle(p, n, title, r, failures)
                results.append(r)
                if failures:
                    break
        except Exception as exc:  # noqa: BLE001 - the run is reverted below
            failures.append(f"{type(exc).__name__}: {com_err(exc)}")
        if failures:
            revert(p)
            out["failed"].append({"slides": run, "titles": [t for _, t in new], "failures": failures,
                                  "reverted": "the whole run (no partial series is kept)"})
            continue
        session.save(p)
        for r in results:
            out["changed"].append({"slide": r["slide"], "title": r["after"], "before": r["before"],
                                   **{k: r[k] for k in ("warnings", "formatting_reset") if k in r}})
    left = title_report(session.find_open(path))["duplicates"]
    if left:
        out["still_duplicate"] = left
    return out
