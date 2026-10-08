"""ppt-a11y: MCP server that drives desktop PowerPoint (Windows, COM) to make decks meet WCAG 2.1 AA:
alt text, slide titles, reading order, text contrast, tables and link text, keeping content, look and
animations.

    python -m ppta11y          (stdio transport; same as server.py at the repository root)
"""
from __future__ import annotations

import functools
import inspect
import json
import sys
import traceback
from pathlib import Path

from mcp.server.mcpserver import Image, MCPServer

from . import animations as anim
from . import audit as audit_mod
from . import checker, inventory, ops, plan, playback, render, session, sheet, structure
from . import contrast, links, readingorder, tables, titles, wcag
from .com import COM, ToolError, com_err

server = MCPServer(
    name="ppt-a11y",
    instructions=(
        "WCAG 2.1 AA / Section 508 remediation of PowerPoint decks through desktop PowerPoint on Windows: alt "
        "text, slide titles, reading order, text contrast, tables and link text, keeping content, look and "
        "animations. Always start with open_working_copy(source): edits are only allowed on that copy. Alt text "
        "per figure: inspect_slide + slide_snapshot -> check_grouping -> group_shapes / set_alt_text / "
        "mark_decorative (group_shapes KEEPS animations, verified, rolls back on any difference). Then: "
        "title_report -> promote_to_title / make_titles_unique; table_report -> table_to_shapes / unmerge_table / "
        "set_table_header; link_report -> set_link_text; contrast_report -> fix_contrast; LAST "
        "reading_order_report + reading_order_snapshot -> fix_reading_order (restacks invisibly, verified). Finish "
        "with wcag_audit (our rules + PowerPoint's own checker): done=true means nothing is left. Shape ids are "
        "PowerPoint Shape.Id values. part='layout'/'master' edits objects inherited from Slide Master view."
    ),
)


def _dump(obj) -> str:
    return json.dumps(obj, indent=1, ensure_ascii=False, default=str)


def tool(fn):
    """Register `fn` as a tool whose body runs on the COM thread; errors come back as text."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return COM.run(fn, *args, **kwargs)
        except ToolError as exc:
            return f"error: {exc}"
        except Exception as exc:  # COM errors etc.
            tb = traceback.format_exc(limit=3)
            return f"error: {type(exc).__name__}: {com_err(exc)}\n{tb[-1200:]}"
    wrapper.__signature__ = inspect.signature(fn)
    return server.tool(structured_output=False)(wrapper)


# --------------------------------------------------------------------------------------------------
# Decks
# --------------------------------------------------------------------------------------------------

@tool
def open_working_copy(source: str, dest: str | None = None, fresh: bool = False) -> str:
    """Copy `source` (.pptx) to a working copy (default '<name>_accessible.pptx' beside it) and open it
    in PowerPoint. All edits go to the working copy; the source is registered as protected and is never
    saved. If the working copy already exists it is reopened (resume), unless fresh=True."""
    return _dump(session.open_working_copy(source, dest, fresh))


@tool
def open_deck(path: str, read_only: bool = True) -> str:
    """Open another deck (e.g. the untouched original for comparison). Read-only by default."""
    return _dump(session.open_deck(path, read_only))


@tool
def list_open_decks() -> str:
    """Presentations open in PowerPoint, and which are working copies / protected originals."""
    return _dump(session.list_open())


@tool
def save_deck(deck: str | None = None) -> str:
    """Save the working copy (refuses anything that is not a registered working copy)."""
    p = session.get_pres(deck)
    session.save(p)
    return _dump({"saved": p.FullName})


@tool
def revert_to_saved(deck: str | None = None) -> str:
    """Discard unsaved changes in the working copy by reopening it from disk."""
    p = session.revert(session.get_pres(deck))
    return _dump({"reverted": p.FullName})


@tool
def close_deck(deck: str | None = None, save: bool = False, discard: bool = False) -> str:
    """Close a deck. Unsaved changes are never dropped silently: a working copy must be saved (save=True)
    or reverted first; any other deck with unsaved changes needs discard=True (ask the user first)."""
    p = session.get_pres(deck)
    name = p.FullName
    session.close(p, save_first=save, discard=discard)
    return _dump({"closed": name})


# --------------------------------------------------------------------------------------------------
# Looking
# --------------------------------------------------------------------------------------------------

@tool
def deck_overview(deck: str | None = None) -> str:
    """One line per slide: title, layout, objects needing alt text, groups, animation count."""
    p = session.get_pres(deck)
    rows = []
    for n in range(1, p.Slides.Count + 1):
        sl = p.Slides(n)
        targets, _ = inventory.timeline_targets(sl)
        need = []
        for i in range(1, sl.Shapes.Count + 1):
            s = sl.Shapes(i)
            st = audit_mod.classify(s)["status"]
            if st in ("missing", "review", "junk"):
                need.append(f"#{s.Id} {st}")
        rows.append({"slide": n, "title": inventory.slide_title(sl), "layout": sl.CustomLayout.Name,
                     "shapes": sl.Shapes.Count, "effects": sum(targets.values()), "todo": need})
    return _dump({"deck": p.FullName, "slides": rows})


@tool
def inspect_slide(slide: int, part: str = "slide", inherited: bool = True, deck: str | None = None) -> str:
    """Every shape on a slide (or its layout/master with part='layout'/'master'): id, kind, alt text,
    decorative flag, text, position (points), stacking order (z), animation effects, group members,
    plus objects inherited from the layout/master. needs_alt: required | review | no."""
    return _dump(inventory.inspect(session.get_pres(deck), slide, part, inherited))


@tool
def slide_snapshot(slide: int, part: str = "slide", ids: list[int] | None = None, nested: bool = False,
                   width: int = 1400, deck: str | None = None) -> list:
    """PowerPoint's own render of the slide with a labelled box per shape ('#id', 'A' = animated).
    Colours: red = needs alt text, orange = shape without text to review, green = has alt text,
    grey = decorative, blue = text. `ids` boxes only those; nested=True also boxes shapes inside groups;
    part='layout'/'master' boxes the inherited objects."""
    p = session.get_pres(deck)
    out = session.out_dir(p) / f"snapshot_slide{slide}_{part}_{session.stamp()}.png"
    path, chosen = render.labeled(p, slide, out, part, ids, nested, width)
    legend = "; ".join(f"#{d['id']} {d['kind']} [{d['status']}]" + (f" '{d['text'][:30]}'" if d.get("text") else "")
                       + (" animated" if d.get("animation_effects") else "") for d in chosen)
    return [f"{inventory.part_label(p, slide, part)} -> {path}\n{legend}", Image(path=str(path))]


@tool
def list_animations(slide: int, deck: str | None = None) -> str:
    """The slide's animation sequence(s) in order: target shape, effect, start (on click / with previous /
    after previous / trigger), delay, duration."""
    p = session.get_pres(deck)
    return _dump(anim.summarize(anim.snapshot(p.Slides(slide))))


# --------------------------------------------------------------------------------------------------
# Editing
# --------------------------------------------------------------------------------------------------

@tool
def check_grouping(slide: int, ids: list[int], part: str = "slide", deck: str | None = None) -> str:
    """Dry run for group_shapes: blockers (placeholders, tables, shapes already in a group), animation
    effects that will be rebuilt, stacking-order warnings. Changes nothing."""
    return _dump(ops.check_group(session.get_pres(deck), slide, ids, part))


@tool
def group_shapes(slide: int, ids: list[int], alt_text: str | None = None, decorative: bool = False,
                 name: str | None = None, part: str = "slide", allow_visual_change: bool = False,
                 deck: str | None = None) -> list:
    """Group top-level shapes (like Ctrl+G) and put `alt_text` (or decorative=True) on the group, KEEPING
    all animation effects: lost effects are re-added to the same shapes inside the group with every
    property and the original order. Verified against PowerPoint's saved animation XML, shape positions
    and a before/after render. Any difference -> automatic rollback to the last save, ok=false.
    On success the working copy is saved."""
    p = session.get_pres(deck)
    res = ops.group_shapes(p, slide, ids, alt_text, decorative, name, part, allow_visual_change)
    img = (res.get("visual_check") or {}).get("image")
    out = [_dump(res)]
    if img:
        out.append(Image(path=img))
    return out


@tool
def set_alt_text(slide: int, ids: list[int], text: str, part: str = "slide", deck: str | None = None) -> str:
    """Set alt text (Shape.AlternativeText) on shapes; clears the decorative flag if set. Returns lint
    warnings (file names, URLs, 'image of...', very long text). Not saved until save_deck."""
    return _dump(ops.set_alt_text(session.get_pres(deck), slide, ids, text, part))


@tool
def mark_decorative(slide: int, ids: list[int], part: str = "slide", decorative: bool = True,
                    deck: str | None = None) -> str:
    """Tick (or untick) 'Mark as decorative' and clear any alt text on those shapes. Not saved until
    save_deck."""
    return _dump(ops.mark_decorative(session.get_pres(deck), slide, ids, part, decorative))


@tool
def clear_alt_text(slide: int, ids: list[int], part: str = "slide", deck: str | None = None) -> str:
    """Remove junk alt text/titles (e.g. 'preencoded.png' or a URL) from shapes."""
    return _dump(ops.clear_alt_text(session.get_pres(deck), slide, ids, part))


@tool
def restore_animations_from(slide: int, source_deck: str, source_slide: int | None = None,
                            replace: bool = False, restore_positions: bool = False, deck: str | None = None) -> list:
    """Re-create the animation sequence of `source_deck` slide `source_slide` (default: same number) on the
    same shape ids of this slide, e.g. after someone grouped a diagram by hand and PowerPoint dropped its
    animations. Verified against the source slide's saved XML; rolls back on mismatch.
    restore_positions=True also moves a group back if grouping shifted it."""
    p = session.get_pres(deck)
    res = ops.restore_animations_from(p, slide, source_deck, source_slide, replace, restore_positions)
    out = [_dump(res)]
    img = (res.get("visual_check") or {}).get("image")
    if img:
        out.append(Image(path=img))
    return out


@tool
def flatten_group(slide: int, group_id: int, allow_alt_loss: bool = False, deck: str | None = None) -> list:
    """Ungroup an existing group that is animated as a whole (Appear/Fade only) and give every piece the same
    effect, starting together, so the pieces can be grouped into the bigger figure (check_grouping recommends
    this when it blocks an animated nested group). Verified: identical viewer timeline + identical render;
    rolls back otherwise."""
    res = structure.flatten_group(session.get_pres(deck), slide, group_id, allow_alt_loss)
    img = (res.get("visual_check") or {}).get("image")
    return [_dump(res)] + ([Image(path=img)] if img else [])


@tool
def placeholder_to_picture(slide: int, shape_id: int, deck: str | None = None) -> list:
    """Replace a picture held in a content/picture placeholder with an identical ordinary picture (position,
    size, crop, stacking, alt text, animations kept) so it can be grouped. Returns new_id (use '@<old id>' in
    apply_plan). Verified render + animation XML; rolls back otherwise. Uses the clipboard."""
    pres = session.get_pres(deck)
    res = structure.placeholder_to_picture(pres, slide, shape_id)
    if res.get("ok") and res.get("new_id"):
        res.update(_remember_ids(deck, {slide: {shape_id: res["new_id"]}}, "slide"))
    img = (res.get("visual_check") or {}).get("image")
    return [_dump(res)] + ([Image(path=img)] if img else [])


@tool
def move_to_slides(slide: int, ids: list[int], part: str = "layout", deck: str | None = None) -> str:
    """Move objects from the layout (or master) used by `slide` onto every slide that shows them, at the same
    place and behind the slide's own content, then delete them from the layout/master. Use when informative
    content (e.g. agenda tiles, a logo with alt text) lives in Slide Master view, where screen readers and the
    checker do not see it. Returns {slide: {old id: new id}}. Verified: every affected slide renders identically.
    Uses the clipboard."""
    res = structure.move_to_slides(session.get_pres(deck), slide, ids, part)
    if res.get("ok") and res.get("mapping"):
        res.update(_remember_ids(deck, res["mapping"], part))
    return _dump(res)


def _remember_ids(deck, mapping: dict, part: str) -> dict:
    """Record {slide: {old id: new id}} in the id map, as apply_plan does, so '@<old id>' works in later
    plans. The edit itself is already saved: a mapping problem is reported, not raised."""
    try:
        p = session.get_pres(deck)
        m = plan.load_map(p)
        for n, mp in mapping.items():
            for old, new in mp.items():
                plan._put(m, int(n), part, int(old), int(new))
        plan._save_map(p, m)
        return {}
    except Exception as exc:  # noqa: BLE001
        return {"id_map_warning": f"saved, but '@<old id>' was not recorded ({com_err(exc)}); use the new ids"}


@tool
def apply_plan(steps: list[dict] | None = None, plan_path: str | None = None, rows: list[int] | None = None,
               stop_on_failure: bool = True, sheet_path: str | None = None, deck: str | None = None) -> str:
    """Run plan steps in order, each a verified transaction (see ppta11y/plan.py for the step format: ops
    group | alt | decorative | clear_alt | flatten_group | placeholder_to_picture | move_to_slides; ids may be
    '@<old id>' for shapes re-created by an earlier move/convert step). `rows` runs only steps with those
    spreadsheet rows. "alt_text": "ROW" / "text": "ROW" copies that row's Alt Text from `sheet_path` verbatim.
    Long plans: run in chunks (about 10-20 s per grouping)."""
    if plan_path:
        steps = json.loads(Path(plan_path).read_text(encoding="utf-8-sig"))
    if not steps:
        raise ToolError("give steps or plan_path")
    if rows:
        steps = plan.select_rows(steps, rows)
    p = session.get_pres(deck)
    return _dump(plan.apply_plan(p.FullName, steps, stop_on_failure, sheet_path))


@tool
def compare_playback(slide: int, source_deck: str, source_slide: int | None = None, deck: str | None = None) -> list:
    """End-to-end animation check: plays the slide of the untouched original and of the working copy in a
    windowed Slide Show, captures a frame after every click and compares them (what the audience sees).
    Run it on every slide where animated shapes were grouped/flattened/converted. Needs the screen
    (PowerPoint window visible); ~1-3 s per click. ok=false lists the frames that differ."""
    p = session.get_pres(deck)
    res = playback.compare_playback(p, slide, source_deck, source_slide)
    if session.is_working_copy(p.FullName) and p.Saved:
        # Slide Show rebuilds PowerPoint's in-memory text layout; reload from disk so later render checks
        # compare against the file, not a stale layout cache.
        session.revert(p)
        res["note"] = "working copy reloaded from disk after playback (clears Slide Show layout cache)"
    return [_dump(res)] + ([Image(path=res["image"])] if res.get("image") else [])


@tool
def lint_alt_text(text: str) -> str:
    """Quality warnings for a proposed alt text."""
    return _dump(ops.lint_alt(text))


# --------------------------------------------------------------------------------------------------
# Checking
# --------------------------------------------------------------------------------------------------

@tool
def audit_alt_text(include_inherited: bool = True, include_ok: bool = False, deck: str | None = None) -> str:
    """Deck-wide audit: every object that still needs alt text ('missing'), shapes without text to review
    ('review'), junk alt text ('junk'), tables without header rows, slides without titles, and the same for
    objects on layouts/master. Writes JSON + CSV reports next to the deck."""
    p = session.get_pres(deck)
    rep = audit_mod.audit(p, include_inherited, include_ok)
    files = audit_mod.write_report(p, rep)
    return _dump({"report_files": files, **rep})


@tool
def run_accessibility_checker(itemize: list[str] | None = None, deck: str | None = None) -> str:
    """Run PowerPoint's own Accessibility Assistant (Review > Check Accessibility) and return its counts per
    category (Missing alt text, Missing slide title, Check reading order, ...). itemize=['Missing alt text']
    steps through each issue and reports the slide + shape PowerPoint selects (about 1 s per issue).
    Note: PowerPoint's checker does not look at slide layouts/master."""
    return _dump(checker.run(session.get_pres(deck), itemize or []))


# --------------------------------------------------------------------------------------------------
# Reading order (WCAG 1.3.2 / 2.4.3)
# --------------------------------------------------------------------------------------------------

@tool
def reading_order_report(slides: list[int] | None = None, compare_to: str | None = None, deck: str | None = None) -> str:
    """What a screen reader reads on each slide, in order (= stacking order, back to front; decorative and
    hidden objects skipped; a group is one item), the suggested order (title, subtitle, content in visual
    reading order with call-out notes right after what they point at, footer/slide number last), issues, and
    with compare_to=<original deck> what earlier edits changed. Slides whose order was set and verified with
    set_reading_order/fix_reading_order count as done. Writes JSON + text reports to the checks folder."""
    return _dump(readingorder.deck_report(session.get_pres(deck), slides, compare_to))


@tool
def reading_order_snapshot(slide: int, suggested: bool = False, deck: str | None = None) -> list:
    """The slide with a numbered badge on every object in the order a screen reader reads it now
    (suggested=True: the suggested order). Green = title, blue = content, grey = footer/slide number."""
    p = session.get_pres(deck)
    items = readingorder.slide_items(p, slide)
    seq = readingorder.suggest(items) if suggested else [it for it in items if it["read"]]
    colour = {"title": (20, 140, 60), "footer": (130, 130, 130)}
    entries = [(i, it["box"], colour.get(it["role"], (40, 90, 200))) for i, it in enumerate(seq, 1)]
    out = session.out_dir(p) / f"ro_s{slide}_{'sug' if suggested else 'now'}_{session.stamp()}.png"
    render.numbered(p, slide, out, entries)
    legend = "\n".join(f"{i}. #{it['id']} {readingorder.spoken(it)}" for i, it in enumerate(seq, 1))
    return [f"slide {slide} reading order ({'suggested' if suggested else 'now'}):\n{legend}", Image(path=str(out))]


@tool
def set_reading_order(slide: int, order: list[int], strict: bool = False, deck: str | None = None) -> str:
    """Make the screen-reader order of a slide `order` (every non-decorative, visible top-level id exactly
    once) by restacking, with NO visible change: overlapping objects keep their stacking (if that prevents
    part of the order, the closest invisible order is used and reported as deviation + not_accepted: check
    reads_now and call again with exactly that order_now to accept it; strict=True fails instead).
    Verified: render, positions, animation XML; rolls back otherwise. Saves on success."""
    return _dump(readingorder.set_reading_order(session.get_pres(deck), slide, order, strict=strict))


@tool
def fix_reading_order(slides: list[int] | None = None, orders: dict | None = None, use_checker: bool = True,
                      hidden_text: str = "decorative", deck: str | None = None) -> str:
    """Apply the suggested reading order to every slide (or `orders` = {"<slide>": [ids]} you chose for
    meaning), each a verified invisible restack. Text nobody can see (completely behind opaque objects, e.g. a
    footer under a full-slide table) is first marked decorative (hidden_text="keep" to leave it). use_checker:
    then run PowerPoint's Accessibility Assistant and, for slides it still flags under "Check reading order",
    try the other suggestion variants (pure visual, placeholders first, top-down, minimal) and keep one it
    accepts."""
    return _dump(readingorder.fix_reading_order(session.get_pres(deck), slides, orders, use_checker,
                                                hidden_text=hidden_text))


# --------------------------------------------------------------------------------------------------
# Slide titles (2.4.2 / 2.4.6; PowerPoint checker: Missing slide title, Duplicate slide title)
# --------------------------------------------------------------------------------------------------

@tool
def title_report(deck: str | None = None) -> str:
    """Slides without a (non-empty) title with the most title-like visible text on each (candidate id), and
    groups of slides that share a title."""
    return _dump(titles.title_report(session.get_pres(deck)))


@tool
def promote_to_title(slide: int, shape_id: int, deck: str | None = None) -> str:
    """Turn the slide's visible heading (text box or text placeholder `shape_id`) into its real title
    placeholder, same position and formatting (looks identical, verified by render). If the layout has no
    title placeholder one is added to it (invisible on every slide, verified). Rolls back otherwise."""
    return _dump(titles.promote_to_title(session.get_pres(deck), slide, shape_id))


@tool
def add_hidden_title(slide: int, text: str, deck: str | None = None) -> str:
    """Fallback when nothing visible can become the title: a title placeholder placed just above the slide
    (Microsoft's "Add Hidden Slide Title"; not shown in Slide Show, read by screen readers). Verified."""
    return _dump(titles.add_hidden_title(session.get_pres(deck), slide, text))


@tool
def set_slide_title(slide: int, text: str, deck: str | None = None) -> str:
    """Change a slide's title text keeping its formatting (visible change, limited to the title; `warnings`
    if the new title wraps, grows or overflows its box). Refuses a title placeholder that is empty, hidden or
    decorative (use promote_to_title / add_hidden_title / mark_decorative(decorative=False))."""
    return _dump(titles.set_slide_title(session.get_pres(deck), slide, text))


@tool
def make_titles_unique(slides: list[int] | None = None, pattern: str = "{title} ({n} of {N})",
                       deck: str | None = None) -> str:
    """Give every run of consecutive slides that share a title a unique one: "{title} ({n} of {N})" by
    default (Section 508 / HHS pattern). Whole runs only (naming one slide numbers its run), one verified
    transaction per run. Non-consecutive duplicates come back in needs_distinct_title. Visible change to
    those titles only."""
    return _dump(titles.make_titles_unique(session.get_pres(deck), slides, pattern))


# --------------------------------------------------------------------------------------------------
# Text contrast (1.4.3)
# --------------------------------------------------------------------------------------------------

@tool
def contrast_report(slides: list[int] | None = None, deck: str | None = None) -> str:
    """Every run of text vs the background it is drawn on (cell/shape fill, else the rendered slide):
    WCAG ratio, required minimum (4.5:1; 3:1 for >= 18 pt or >= 14 pt bold, rendered size), pass/fail.
    Finds failures PowerPoint's checker misses (text on the slide background, inside groups, over other
    shapes, layout/master text, text under/over animated objects in both states, dimmed text and colour
    emphasis states as rows with "state")."""
    return _dump(contrast.contrast_report(session.get_pres(deck), slides))


@tool
def fix_contrast(slides: list[int] | None = None, strategy: str = "auto", deck: str | None = None) -> str:
    """Make every failing run pass with the smallest visible change: the nearest passing text colour (same
    hue), or (strategy="auto") a slightly darker/lighter cell/shape fill when that changes much less (e.g.
    white text on a mid-tone fill). strategy="text" never changes fills. Re-measured; one save per slide.
    Layout/master text gets one colour that passes on every slide showing it. Colours stored inside an
    animation (dim, colour emphasis) are returned in `unfixable` with the colour to pick by hand."""
    return _dump(contrast.fix_contrast(session.get_pres(deck), slides, strategy))


# --------------------------------------------------------------------------------------------------
# Tables (1.3.1; PowerPoint checker: Missing table header, Use of merged or split cells)
# --------------------------------------------------------------------------------------------------

@tool
def table_report(deck: str | None = None) -> str:
    """Every table: size, merged regions, header row, whether it is a hierarchy/layout table, recommended fix."""
    return _dump(tables.table_report(session.get_pres(deck)))


@tool
def table_to_shapes(slide: int, table_id: int, reading: str = "auto", deck: str | None = None) -> str:
    """Layout/diagram tables drawn with merged cells (trees, roadmaps): replace the table with real text boxes
    (same geometry, fills, formatted text) plus its borders as one decorative group, read branch by branch
    (reading: auto | depth_first | row_major). Looks the same (verified allowing 1-px antialiasing)."""
    return _dump(tables.table_to_shapes(session.get_pres(deck), slide, table_id, reading))


@tool
def unmerge_table(slide: int, table_id: int, repeat_text: bool = True, deck: str | None = None) -> str:
    """Data tables: split merged regions back into single cells (grid unchanged) and repeat the merged text in
    each cell (Section 508 simple-table pattern). Visible change, reported with a compare image."""
    return _dump(tables.unmerge_table(session.get_pres(deck), slide, table_id, repeat_text))


@tool
def set_table_header(slide: int, table_id: int, deck: str | None = None) -> str:
    """Turn on Header Row without restyling the table (fills and text colours re-applied; verified)."""
    return _dump(tables.set_table_header(session.get_pres(deck), slide, table_id))


# --------------------------------------------------------------------------------------------------
# Link text (2.4.4)
# --------------------------------------------------------------------------------------------------

@tool
def link_report(deck: str | None = None) -> str:
    """Every text hyperlink: where, text, address; `unclear` (fails 2.4.4 / flagged by PowerPoint: raw URL alone
    in its paragraph, long URL with a query string, generic text, or no target) with a suggested text, and
    `advice` for links that pass but could be clearer."""
    return _dump(links.link_report(session.get_pres(deck)))


@tool
def set_link_text(slide: int, shape_id: int, old_text: str, new_text: str, address: str | None = None,
                  cell: list[int] | None = None, deck: str | None = None) -> str:
    """Replace a link's visible text with descriptive text (use the destination page's title) and keep the
    link (a dead link whose text was its URL gets that URL as address). Only those characters change."""
    return _dump(links.set_link_text(session.get_pres(deck), slide, shape_id, old_text, new_text, address, cell))


# --------------------------------------------------------------------------------------------------
# Whole deck
# --------------------------------------------------------------------------------------------------

@tool
def wcag_audit(compare_to: str | None = None, run_checker: bool = True, check_contrast: bool = True,
               deck: str | None = None) -> str:
    """Everything WCAG 2.1 AA needs on this deck, with counts of what is left: alt text, titles (missing,
    duplicate), reading order, contrast, tables (merged, header), links; plus PowerPoint's own checker.
    done=true means nothing is left. Writes a JSON report to the checks folder."""
    return _dump(wcag.wcag_audit(session.get_pres(deck), compare_to, run_checker, check_contrast))


# --------------------------------------------------------------------------------------------------
# Spreadsheet
# --------------------------------------------------------------------------------------------------

@tool
def read_alt_text_sheet(path: str, sheet_name: str | None = None) -> str:
    """Rows of an alt-text spreadsheet (header row = column names), with Excel row numbers."""
    return _dump(sheet.read(path, sheet_name))


@tool
def record_change_log(dest: str, changes: list[dict], sheet_name: str = "WCAG changes", replace: bool = True) -> str:
    """Write a WCAG change log to its own sheet of the results copy `dest` (made by record_sheet_results, or a
    new workbook): one row per change with Slide, Category, WCAG, What changed, Kind (invisible/visible),
    Before, After, Verified. The original spreadsheet is never touched. replace=True (default) rewrites the
    sheet on every call: pass all changes at once (replace=False appends). Only a change-log sheet is ever
    rewritten."""
    return _dump(sheet.record_changes(dest, changes, sheet_name, replace))


@tool
def record_sheet_results(source: str, dest: str, updates: dict, sheet_name: str | None = None) -> str:
    """Write results into new columns of a COPY of the spreadsheet (`dest`, created from `source` on first
    use). updates = {"<excel row>": {"<column header>": "<value>"}}. Input columns are never edited."""
    return _dump(sheet.record(source, dest, updates, sheet_name))


# --------------------------------------------------------------------------------------------------
# Escape hatch
# --------------------------------------------------------------------------------------------------

@tool
def evaluate(code: str, slide: int | None = None, deck: str | None = None) -> str:
    """Run Python against the deck for inspection the other tools do not cover. Globals: presentation,
    slide (if `slide` given), find(container, id). Set `result` to return data. Only allowed on a working
    copy or a deck opened read-only. This is NOT a sandbox: never save, close or edit other decks from here
    (use the dedicated tools); the call reports an error if any other open deck was changed."""
    p = session.get_pres(deck)
    if not (session.is_working_copy(p.FullName) or p.ReadOnly):
        raise ToolError("evaluate is only allowed on a working copy or a read-only deck")
    if any(w in code for w in (".SaveAs", ".Save(", ".SaveCopyAs", ".Close(", "Presentations.Open")):
        raise ToolError("saving/closing/opening decks from evaluate is not allowed; use the dedicated tools")
    others = {o.FullName: bool(o.Saved) for o in p.Application.Presentations if o.FullName != p.FullName}
    g = {"presentation": p, "slide": p.Slides(slide) if slide else None,
         "find": inventory.find_shape, "result": None}
    exec(code, g)
    after = {o.FullName: bool(o.Saved) for o in p.Application.Presentations if o.FullName != p.FullName}
    touched = [n for n, saved in others.items() if after.get(n) != saved]
    if touched:
        raise ToolError(f"evaluate changed other open decks {touched}; review them in PowerPoint (nothing saved)")
    return _dump(g.get("result"))


def main() -> None:
    """Run the MCP server over stdio."""
    server.run()


if __name__ == "__main__":
    main()
