"""Tables. PowerPoint cannot express header relationships for merged or split cells, so Section 508 fails
any table with merged cells, and PowerPoint's checker warns about "Use of merged or split cells"
(its own heuristic only fires on tables of 4x4 and larger). Every table also needs a header row
("Missing table header").

  table_report     every table: size, merged regions, header row, whether it is a layout/diagram table
  table_to_shapes  a layout or diagram table drawn with merged cells (a tree, a roadmap) becomes real text
                   boxes with the same geometry, fills, text and borders (borders as one decorative group),
                   read in a sensible order (branch by branch for trees); looks the same, verified by render
  unmerge_table    a data table: every merged region is split back into single cells and the merged text
                   repeated in each (Section 508 / HHS); visible change, reported
  set_table_header turn on Header Row, keeping every cell's fill and text colour as they were
"""
from __future__ import annotations

import re

from . import render
from .com import ToolError, com_err, safe
from .inventory import get_part, timeline_targets, top_level_index
from .session import checkpoint, out_dir, require_editable, revert, save, stamp
from .textfmt import apply_color_state, color_state, copy_color, copy_text

TOL = 1.0
MIXED = -2147483648


def _table_shape(sl, table_id: int):
    i = top_level_index(sl, table_id)
    if i is None:
        raise ToolError(f"#{table_id} is not a top-level shape")
    s = sl.Shapes(i)
    if safe(lambda: s.HasTable) != -1:
        raise ToolError(f"#{table_id} is not a table")
    return s


def grid(s) -> dict:
    """Grid lines and merged regions of table shape `s` (regions in reading-grid order)."""
    t = s.Table
    R, C = t.Rows.Count, t.Columns.Count
    xs, ys = [s.Left], [s.Top]
    for c in range(1, C + 1):
        xs.append(xs[-1] + t.Columns(c).Width)
    for r in range(1, R + 1):
        ys.append(ys[-1] + t.Rows(r).Height)
    owner, regions = {}, []
    for r in range(1, R + 1):
        for c in range(1, C + 1):
            if (r, c) in owner:
                continue
            cs = t.Cell(r, c).Shape
            right, bottom = cs.Left + cs.Width, cs.Top + cs.Height
            c1 = c
            while c1 < C and xs[c1] < right - TOL:
                c1 += 1
            r1 = r
            while r1 < R and ys[r1] < bottom - TOL:
                r1 += 1
            k = len(regions)
            regions.append({"r0": r, "c0": c, "r1": r1, "c1": c1, "box": (xs[c - 1], ys[r - 1], xs[c1], ys[r1]),
                            "text": (cs.TextFrame.TextRange.Text or "").replace("\r", " / ")[:80]})
            for rr in range(r, r1 + 1):
                for cc in range(c, c1 + 1):
                    owner[(rr, cc)] = k
    merged = [g for g in regions if g["r1"] > g["r0"] or g["c1"] > g["c0"]]
    return {"rows": R, "cols": C, "xs": xs, "ys": ys, "regions": regions, "owner": owner, "merged": merged}


def _tree(regions) -> dict | None:
    """{region index: [children]} for a hierarchy drawn with merged cells (a tree / roadmap), else None.
    Requires: every region below row 1 sits under exactly one region of the row just above it that spans
    it; at least one region branches into two or more; no region spans rows (row spans are data-table row
    groups); and at most one row - the last - is made only of single cells (more such rows = a data grid
    with a merged header, which is a data table, not a diagram)."""
    if not regions or any(r["r1"] > r["r0"] for r in regions):
        return None
    kids = {i: [] for i in range(len(regions))}
    for j, b in enumerate(regions):
        if b["r0"] == 1:
            continue
        parents = [i for i, a in enumerate(regions)
                   if a["r1"] == b["r0"] - 1 and a["c0"] <= b["c0"] and b["c1"] <= a["c1"]]
        if len(parents) != 1:
            return None
        kids[parents[0]].append(j)
    if not any(len(v) >= 2 for v in kids.values()):
        return None
    last = max(r["r0"] for r in regions)
    single_rows = {r["r0"] for r in regions} - {r["r0"] for r in regions if r["c1"] > r["c0"]}
    if single_rows - {last}:
        return None
    # a column of 3+ single cells one under another is a data column (an item list), not a branch
    singles = {(r["r0"], r["c0"]) for r in regions if r["c1"] == r["c0"]}
    if any((r0 + 1, c) in singles and (r0 + 2, c) in singles for r0, c in singles):
        return None
    leaves = [regions[i]["text"].strip() for i, v in kids.items() if not v and regions[i]["text"].strip()]
    if leaves and sum(1 for x in leaves if _NUM.match(x.replace(" ", ""))) >= 0.5 * len(leaves):
        return None  # numbers under headings: a data table with a spanning header
    for v in kids.values():
        v.sort(key=lambda j: regions[j]["c0"])
    return kids


def reading_sequence(g: dict, mode: str = "auto") -> tuple[list[int], str]:
    regions = g["regions"]
    tree = _tree(regions) if mode in ("auto", "depth_first") else None
    if tree is not None and g["merged"]:
        out = []

        def dfs(i):
            out.append(i)
            for j in tree[i]:
                dfs(j)
        for i in sorted([i for i, r in enumerate(regions) if r["r0"] == 1], key=lambda i: regions[i]["c0"]):
            dfs(i)
        return out, "branch by branch (a hierarchy: each heading, then everything under it)"
    if mode == "depth_first":
        raise ToolError("this table is not a hierarchy; use reading='row_major'")
    return sorted(range(len(regions)), key=lambda i: (regions[i]["r0"], regions[i]["c0"])), "row by row"


_NUM = re.compile(r"^[-+(]?[$€£]?\d[\d,.]*%?\)?$")


def _multi_row_header(s, g) -> bool:
    """Row 2 holds column labels under a spanning/repeated row 1, over numeric body rows: a two-row header,
    which PowerPoint cannot mark up (only row 1 is announced as headers)."""
    t, R, C = s.Table, g["rows"], g["cols"]
    if R < 3:
        return False
    txt = lambda r, c: (safe(lambda: t.Cell(r, c).Shape.TextFrame.TextRange.Text, "") or "").strip()  # noqa: E731
    row1 = [txt(1, c) for c in range(1, C + 1)]
    spans = any(r["r0"] == 1 and r["c1"] > r["c0"] for r in g["regions"]) or \
        any(row1[c] and row1[c] == row1[c + 1] for c in range(C - 1))
    row2 = [txt(2, c) for c in range(1, C + 1)]
    labels = sum(1 for x in row2 if x and not _NUM.match(x.replace(" ", ""))) >= 0.6 * C
    body = [txt(r, c) for r in range(3, R + 1) for c in range(1, C + 1)]
    body = [x for x in body if x]
    numeric = body and sum(1 for x in body if _NUM.match(x.replace(" ", ""))) >= 0.5 * len(body)
    return bool(spans and labels and numeric)


def table_report(pres) -> dict:
    rows = []
    for n in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(n)
        for i in range(1, sl.Shapes.Count + 1):
            s = sl.Shapes(i)
            if safe(lambda: s.HasTable) != -1:
                continue
            g = grid(s)
            tree = _tree(g["regions"]) is not None and bool(g["merged"])
            header = bool(safe(lambda: s.Table.FirstRow, 0))
            multi = _multi_row_header(s, g)
            rec = []
            if g["merged"]:
                rec.append("table_to_shapes (diagram drawn as a table: looks the same)" if tree else
                           "unmerge_table (data table; visible) - or table_to_shapes only if it is really a diagram")
            if not header:
                rec.append("set_table_header")
            if multi:
                rec.append("two header rows: combine them into one (e.g. 'Sales Q1'), visible change - ask first")
            rows.append({"slide": n, "id": s.Id, "name": s.Name, "rows": g["rows"], "cols": g["cols"],
                         "merged_regions": len(g["merged"]), "header_row": header, "hierarchy": tree,
                         "multi_row_header": multi, "first_cells": [r["text"] for r in g["regions"][:4]],
                         "recommended": rec})
    return {"tables": rows, "summary": {"tables": len(rows),
                                        "with_merged_cells": [(r["slide"], r["id"]) for r in rows if r["merged_regions"]],
                                        "without_header_row": [(r["slide"], r["id"]) for r in rows if not r["header_row"]],
                                        "multi_row_header": [(r["slide"], r["id"]) for r in rows if r["multi_row_header"]]}}


def _border(t, r, c, side):
    b = t.Cell(r, c).Borders(side)
    vis = safe(lambda: b.Visible, 0)
    w = safe(lambda: b.Weight, 0) or 0
    st = color_state(b.ForeColor)
    if vis != -1 or w <= 0 or w == MIXED or st is None or (safe(lambda: b.Transparency, 0) or 0) >= 0.999:
        return None
    return (round(w, 3), st["rgb"], safe(lambda: b.DashStyle, 1) or 1, st["theme"], st["brightness"], st["tint"])


def _segments(s, g) -> list[tuple]:
    """Border segments as (x1, y1, x2, y2, (weight, rgb, dash, theme, brightness, tint)), collinear equal
    pieces merged."""
    t, R, C, xs, ys, own = s.Table, g["rows"], g["cols"], g["xs"], g["ys"], g["owner"]
    raw = []
    for k in range(R + 1):  # horizontal boundary below row k (k=0: top edge)
        for c in range(1, C + 1):
            above, below = (k, c) if k >= 1 else None, (k + 1, c) if k + 1 <= R else None
            if above and below and own[above] == own[below]:
                continue
            spec = (_border(t, *below, 1) if below else None) or (_border(t, *above, 3) if above else None)
            if spec:
                raw.append(("h", ys[k], xs[c - 1], xs[c], spec))
    for k in range(C + 1):  # vertical boundary right of column k
        for r in range(1, R + 1):
            left, right = (r, k) if k >= 1 else None, (r, k + 1) if k + 1 <= C else None
            if left and right and own[left] == own[right]:
                continue
            spec = (_border(t, *right, 2) if right else None) or (_border(t, *left, 4) if left else None)
            if spec:
                raw.append(("v", xs[k], ys[r - 1], ys[r], spec))
    raw.sort(key=lambda x: (x[0], round(x[1], 2), x[2]))
    merged = []
    for seg in raw:
        if merged and merged[-1][0] == seg[0] and abs(merged[-1][1] - seg[1]) < 0.01 and \
                abs(merged[-1][3] - seg[2]) < 0.01 and merged[-1][4] == seg[4]:
            merged[-1] = (seg[0], seg[1], merged[-1][2], seg[3], seg[4])
        else:
            merged.append(seg)
    out = []
    for o, at, a, b, spec in merged:
        ext = spec[0] / 2  # square the corners like table borders
        out.append((a - ext, at, b + ext, at, spec) if o == "h" else (at, a - ext, at, b + ext, spec))
    return out


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


def table_to_shapes(pres, slide: int, table_id: int, reading: str = "auto", max_changed_fraction: float = 0.004,
                    autosave: bool = True) -> dict:
    """Replace a layout/diagram table (merged cells) with real text boxes: one per visible cell, same
    geometry, fill and formatted text, created in reading order (`reading`: auto | depth_first | row_major),
    plus its borders as lines in one decorative group, all at the table's place in the stacking order.
    Verified by render (only antialiasing-level differences allowed). Rolls back otherwise."""
    from .readingorder import _restack
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    s = _table_shape(sl, table_id)
    targets, triggers = timeline_targets(sl)
    if targets.get(table_id) or table_id in triggers:
        raise ToolError(f"table #{table_id} is animated; converting would drop its animation")
    g = grid(s)
    seq, how = reading_sequence(g, reading)
    for r in g["regions"]:
        cs = s.Table.Cell(r["r0"], r["c0"]).Shape
        if safe(lambda: cs.Fill.Visible, 0) == -1 and safe(lambda: cs.Fill.Type) not in (1, None):
            raise ToolError(f"cell R{r['r0']}C{r['c0']} has a gradient/picture fill; not converted")
    checkpoint(pres)
    result = {"slide": slide, "table": table_id, "cells": len(g["regions"]), "merged_regions": len(g["merged"]),
              "reading": how}
    failures, log = [], []
    try:
        od, tag = out_dir(pres), f"s{slide}_t2s_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        z_index = s.ZOrderPosition - 1
        alt = ((s.AlternativeText or "").strip(), (safe(lambda: s.Title, "") or "").strip())
        if any(alt):
            result["table_alt_text_not_kept"] = {"alt_text": alt[0], "title": alt[1],
                                                 "note": "the table's own alt text has no place on text boxes; "
                                                         "put it in the slide notes if it adds information"}
        new_ids, deco_ids = [], []
        for i in seq:
            r = g["regions"][i]
            cs = s.Table.Cell(r["r0"], r["c0"]).Shape
            filled = safe(lambda: cs.Fill.Visible, 0) == -1
            empty = not (cs.TextFrame.TextRange.Text or "").strip()
            if empty and not filled:
                continue  # nothing to draw, nothing to read
            x0, y0, x1, y1 = r["box"]
            box = sl.Shapes.AddShape(1, x0, y0, x1 - x0, y1 - y0)  # msoShapeRectangle
            box.Line.Visible = 0
            safe(lambda: setattr(box.Shadow, "Visible", 0))
            if filled:
                box.Fill.Visible = -1
                box.Fill.Solid()
                copy_color(cs.Fill.ForeColor, box.Fill.ForeColor)  # theme tints/shades kept
                box.Fill.Transparency = safe(lambda: cs.Fill.Transparency, 0) or 0
            else:
                box.Fill.Visible = 0
            copy_text(cs, box, log)
            box.TextFrame2.AutoSize = 0
            box.TextFrame2.WordWrap = -1
            box.Left, box.Top, box.Width, box.Height = x0, y0, x1 - x0, y1 - y0
            box.Name = f"Table cell R{r['r0']}C{r['c0']}"
            if empty:  # a coloured but empty cell: decoration, not a reading stop
                box.Decorative = -1
                deco_ids.append(box.Id)
                continue
            box.Tags.Add("A11Y_BLOCK", f"table{table_id}")   # read as one block, in this order
            box.Tags.Add("A11Y_SEQ", str(len(new_ids) + 1))
            new_ids.append(box.Id)
        lines = []
        for x1, y1, x2, y2, (w, rgb, dash, theme, bright, tint) in _segments(s, g):
            ln = sl.Shapes.AddLine(x1, y1, x2, y2)
            ln.Line.Visible = -1
            ln.Line.Weight = w
            apply_color_state(ln.Line.ForeColor, {"theme": theme, "brightness": bright, "tint": tint, "rgb": rgb})
            safe(lambda: setattr(ln.Line, "DashStyle", dash))
            lines.append(ln.Id)
        grid_id = None
        if len(lines) > 1:
            idx = [top_level_index(sl, i) for i in lines]
            gg = sl.Shapes.Range(idx).Group()
            gg.Name = "Table grid lines"
            gg.Decorative = -1
            grid_id = gg.Id
        elif lines:
            ln = sl.Shapes(top_level_index(sl, lines[0]))
            ln.Decorative = -1
            grid_id = ln.Id
        was_placeholder, box0 = s.Type == 14, (s.Left, s.Top)
        ids_before = {sl.Shapes(i).Id for i in range(1, sl.Shapes.Count + 1)}
        s.Delete()
        if was_placeholder:  # a table in a content placeholder leaves an empty placeholder behind
            from .titles import _remove_leftover_placeholder
            _remove_leftover_placeholder(sl, ids_before, box0, log)
        cur = [sl.Shapes(i).Id for i in range(1, sl.Shapes.Count + 1)]
        mine = deco_ids + new_ids + ([grid_id] if grid_id else [])
        rest = [x for x in cur if x not in mine]
        _restack(sl, rest[:z_index] + mine + rest[z_index:])
        img1 = render.export(pres, slide, od / f"{tag}_after.png")
        vis = render.diff(img0, img1, od / f"{tag}_compare.png")
        # Table borders and lines are antialiased differently (a 1-px edge, a dark-grey vs black pixel), so
        # compare allowing a 1-px shift: small residue allowed at normal sensitivity, none at structural
        # sensitivity (missing/moved text, missing lines and wrong fills all show up there).
        area_px = (g["xs"][-1] - g["xs"][0]) * (g["ys"][-1] - g["ys"][0]) * (render.WIDTH / pres.PageSetup.SlideWidth) ** 2
        allowed = max(40, int(max_changed_fraction * area_px))
        soft = render.tolerant_changed(img0, img1, threshold=24)
        hard = render.tolerant_changed(img0, img1, threshold=72)
        vis.update(shift_tolerant_changed=soft, structural_changed=hard)
        if soft > allowed or hard > 50:
            failures.append(f"the converted table looks different ({soft} px changed allowing 1-px antialiasing, allowed "
                            f"{allowed}; {hard} px structurally; see {vis.get('image')})")
        result.update(text_boxes=new_ids, grid_lines=grid_id, visual_check=vis, allowed_changed_px=allowed,
                      refused_settings=log or None,
                      note=("the table is now real text boxes read " + how + "; borders are one decorative group. "
                            "Merged-cell and header-row checks no longer apply. Run fix_reading_order next."))
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


def unmerge_table(pres, slide: int, table_id: int, repeat_text: bool = True, autosave: bool = True) -> dict:
    """Data tables: split every merged region back into single cells (exact span, so the grid is not
    changed) and repeat the merged text in each of its cells. Visible change (reported with an image)."""
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    s = _table_shape(sl, table_id)
    g = grid(s)
    if not g["merged"]:
        return {"slide": slide, "table": table_id, "ok": True, "unchanged": "no merged cells"}
    from .readingorder import overlaps
    checkpoint(pres)
    result = {"slide": slide, "table": table_id, "regions_split": len(g["merged"])}
    failures, log = [], []
    box_of = lambda x: (x.Left, x.Top, x.Left + x.Width, x.Top + x.Height)  # noqa: E731
    others0 = {sl.Shapes(i).Id: box_of(sl.Shapes(i)) for i in range(1, sl.Shapes.Count + 1)
               if sl.Shapes(i).Id != table_id and not safe(lambda i=i: sl.Shapes(i).PlaceholderFormat.Type) in (13, 15, 16)}
    tb0 = box_of(s)
    try:
        od, tag = out_dir(pres), f"s{slide}_unmerge_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        t = s.Table
        for r in sorted(g["merged"], key=lambda r: (-r["r0"], -r["c0"])):
            t.Cell(r["r0"], r["c0"]).Split(r["r1"] - r["r0"] + 1, r["c1"] - r["c0"] + 1)
            if repeat_text:
                src = t.Cell(r["r0"], r["c0"]).Shape
                for rr in range(r["r0"], r["r1"] + 1):
                    for cc in range(r["c0"], r["c1"] + 1):
                        if (rr, cc) != (r["r0"], r["c0"]):
                            copy_text(src, t.Cell(rr, cc).Shape, log, frame=False)
        s2 = sl.Shapes(top_level_index(sl, table_id))
        g2 = grid(s2)
        if g2["merged"] or (g2["rows"], g2["cols"]) != (g["rows"], g["cols"]):
            failures.append(f"table grid changed or still merged ({g2['rows']}x{g2['cols']}, {len(g2['merged'])} merged)")
        tb1 = box_of(s2)
        result.update(height_before=round(tb0[3] - tb0[1], 1), height_after=round(tb1[3] - tb1[1], 1))
        H = pres.PageSetup.SlideHeight
        if tb1[3] > max(H, tb0[3]) + 0.5:
            failures.append(f"the repeated text makes the table run off the slide (bottom {tb1[3]:.0f} pt > "
                            f"{max(H, tb0[3]):.0f} pt); shorten the merged labels or split the table")
        elif tb0[3] > H + 0.5:
            result["warning"] = (f"the table already ran past the bottom of the slide before (bottom {tb0[3]:.0f} pt > "
                                 f"{H:.0f} pt); unmerging did not make it taller")
        newly = [sid for sid, b in others0.items() if overlaps(tb1, b) and not overlaps(tb0, b)]
        if newly:
            failures.append(f"the taller table now covers shapes {newly}")
        vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
        result.update(visual_check=vis, refused_settings=log or None,
                      note="visible change: merged text now repeated in each cell (Section 508 simple-table pattern)")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)


def _cell_state(t, r, c, FONT) -> tuple:
    cs = t.Cell(r, c).Shape
    tr2 = cs.TextFrame2.TextRange
    runs = tuple((x.Start, x.Length, tuple(sorted((color_state(x.Font.Fill.ForeColor) or {}).items())),
                  tuple((k, safe(lambda k=k: getattr(x.Font, k))) for k in FONT))
                 for x in (tr2.Runs.Item(i) for i in range(1, tr2.Runs.Count + 1)))
    fill = (safe(lambda: cs.Fill.Visible), tuple(sorted((color_state(cs.Fill.ForeColor) or {}).items())),
            safe(lambda: cs.Fill.Transparency, 0))
    return fill, runs


def set_table_header(pres, slide: int, table_id: int, autosave: bool = True) -> dict:
    """Turn on Header Row (what PowerPoint's checker and screen readers need) without restyling the table:
    each cell whose fill, text colour or font the header style changed gets its old values back as direct
    formatting (theme colours with their tints/shades). Verified by render."""
    from .ops import VISUAL_TOLERANCE_PX
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    s = _table_shape(sl, table_id)
    t = s.Table
    if safe(lambda: t.FirstRow, 0):
        return {"slide": slide, "table": table_id, "ok": True, "unchanged": "header row already on"}
    FONT = ("Bold", "Italic", "UnderlineStyle", "Size", "Name", "Allcaps", "Smallcaps")
    keep, borders = {}, {}
    for r in range(1, t.Rows.Count + 1):
        for c in range(1, t.Columns.Count + 1):
            keep[(r, c)] = _cell_state(t, r, c, FONT)
            for side in (1, 2, 3, 4):  # header-row styles add their own borders: put the old ones back
                b = t.Cell(r, c).Borders(side)
                borders[(r, c, side)] = {k: safe(lambda k=k: getattr(b, k))
                                         for k in ("Visible", "Weight", "Transparency", "DashStyle")}
                borders[(r, c, side)]["color"] = color_state(b.ForeColor)
    checkpoint(pres)
    result = {"slide": slide, "table": table_id}
    failures = []
    try:
        od, tag = out_dir(pres), f"s{slide}_thead_{stamp()}"
        img0 = render.export(pres, slide, od / f"{tag}_before.png")
        t.FirstRow = -1
        restored = 0
        for (r, c), (fill, runs) in keep.items():
            if _cell_state(t, r, c, FONT) == (fill, runs):
                continue  # the header style did not touch this cell: leave it alone
            restored += 1
            cs = t.Cell(r, c).Shape
            fv, fst, tr = fill
            if fv == -1 and fst:
                cs.Fill.Visible = -1
                cs.Fill.Solid()
                apply_color_state(cs.Fill.ForeColor, dict(fst))
                cs.Fill.Transparency = tr or 0
            elif fv == 0:
                cs.Fill.Visible = 0
            for start, length, col, font in runs:
                rng = cs.TextFrame2.TextRange.GetCharacters(start, length)
                if col and (color_state(rng.Font.Fill.ForeColor) or {}) != dict(col):
                    apply_color_state(rng.Font.Fill.ForeColor, dict(col))
                for k, v in font:
                    if v is not None and v != MIXED and safe(lambda k=k: getattr(rng.Font, k)) != v:
                        safe(lambda k=k, v=v: setattr(rng.Font, k, v))
        for (r, c, side), b in borders.items():
            bd = t.Cell(r, c).Borders(side)
            now = {k: safe(lambda k=k: getattr(bd, k)) for k in ("Visible", "Weight", "Transparency", "DashStyle")}
            if now == {k: b[k] for k in now} and color_state(bd.ForeColor) == b["color"]:
                continue
            if b["Visible"] == 0 or not b["Weight"] or b["Weight"] == MIXED:
                safe(lambda: setattr(bd, "Visible", 0))
                continue
            for k in ("Visible", "Weight", "Transparency", "DashStyle"):
                if b[k] is not None and b[k] != MIXED:
                    safe(lambda k=k: setattr(bd, k, b[k]))
            if b["color"]:
                safe(lambda: apply_color_state(bd.ForeColor, b["color"]))
        vis = render.diff(img0, render.export(pres, slide, od / f"{tag}_after.png"), od / f"{tag}_compare.png")
        if vis["changed_pixels"] > VISUAL_TOLERANCE_PX:
            failures.append(f"table looks different with Header Row on ({vis['changed_pixels']} px)")
        result.update(visual_check=vis, cells_restored=restored,
                      note="Header Row is on; the first row's cells are announced as column headers")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    return _finish(pres, result, failures, autosave)
