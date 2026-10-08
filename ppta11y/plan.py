"""Run a list of accessibility edits in order, each as its own verified transaction.

Plan = JSON list of steps:
  {"row": 11, "op": "placeholder_to_picture", "slide": 11, "id": 10}
  {"row": 11, "op": "group", "slide": 11, "ids": ["@10", 11, 12], "alt_text": "ROW", "name": "..."}
  {"row": 4,  "op": "move_to_slides", "slide": 2, "part": "layout", "ids": [11, 13, 14]}
  {"row": 4,  "op": "group", "slide": 60, "ids": ["@11", "@13", "@14"], "alt_text": "ROW"}
  {"row": 16, "op": "flatten_group", "slide": 17, "id": 24}
  {"op": "alt" | "decorative" | "clear_alt", "slide": 5, "part": "slide", "ids": [...], "text": "..."}

"@N" = the shape that #N became on this slide after an earlier move_to_slides / placeholder_to_picture.
"@layout:N" / "@master:N" / "@slide:N" pick the source part explicitly when an id exists in several.
The mapping is kept in <deck>_a11y_checks/id_map.json (keys "<slide>:<part>:<old id>"), so a plan can be
run in chunks. Structural steps that were already applied are skipped, so a chunk can be re-run.
"""
from __future__ import annotations

import json
import os

from . import ops, session, structure
from .com import ToolError, com_err, safe
from .inventory import top_level_index

OPS = {"group", "alt", "decorative", "clear_alt", "flatten_group", "placeholder_to_picture", "move_to_slides",
       # WCAG beyond alt text (v2)
       "promote_title", "hidden_title", "set_title", "titles_unique", "table_to_shapes", "unmerge_table",
       "table_header", "link_text", "contrast", "reading_order", "fix_reading_order"}
DECK_OPS = {"titles_unique", "contrast", "fix_reading_order"}  # "slide" optional: they take "slides"
ID_OPS = {"flatten_group", "placeholder_to_picture", "promote_title", "table_to_shapes", "unmerge_table",
          "table_header", "link_text"}
PARTS = ("slide", "layout", "master")


def _map_path(pres):
    return session.out_dir(pres) / "id_map.json"


def load_map(pres) -> dict:
    p = _map_path(pres)
    try:
        return json.loads(p.read_text(encoding="utf8"))
    except FileNotFoundError:
        return {}
    except Exception as exc:  # never silently drop the mapping
        raise ToolError(f"cannot read {p}: {exc}; fix or restore it before continuing")


def _save_map(pres, m: dict) -> None:
    p = _map_path(pres)
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(m, indent=1), encoding="utf8")
    os.replace(tmp, p)


def _put(m: dict, slide: int, part: str, old: int, new: int) -> None:
    k = f"{slide}:{part}:{old}"
    if k in m and int(m[k]) != int(new):
        raise ToolError(f"id map already has {k} -> {m[k]}; refusing to overwrite it with {new}")
    m[k] = int(new)


def _resolve(m: dict, slide: int, ids) -> list[int]:
    out = []
    for x in ids:
        if not (isinstance(x, str) and x.startswith("@")):
            out.append(int(x))
            continue
        ref = x[1:]
        if ":" in ref:
            part, n = ref.split(":", 1)
            keys = [f"{slide}:{part}:{int(n)}"]
        else:
            keys = [f"{slide}:{p}:{int(ref)}" for p in PARTS] + [f"{slide}:{int(ref)}"]  # last = legacy format
        vals = {int(m[k]) for k in keys if k in m}
        if not vals:
            raise ToolError(f"{x} on slide {slide} has no recorded new id (run its move/convert step first)")
        if len(vals) > 1:
            raise ToolError(f"{x} on slide {slide} is ambiguous ({sorted(vals)}); use @slide:/@layout:/@master:")
        out.append(vals.pop())
    return out


def validate(steps: list[dict]) -> None:
    for i, s in enumerate(steps, 1):
        op = s.get("op")
        if op == "note":
            continue
        if op not in OPS:
            raise ToolError(f"step {i}: unknown op {op!r}")
        if "slide" not in s and op not in DECK_OPS:
            raise ToolError(f"step {i}: missing 'slide'")
        if op in ID_OPS and "id" not in s:
            raise ToolError(f"step {i}: {op} needs 'id'")
        if op in ("group", "alt", "decorative", "clear_alt", "move_to_slides", "reading_order") and not s.get("ids"):
            raise ToolError(f"step {i}: {op} needs 'ids'")
        if op in ("alt", "hidden_title", "set_title") and not s.get("text"):
            raise ToolError(f"step {i}: {op} needs 'text'")
        if op == "link_text" and not (s.get("old_text") and s.get("new_text")):
            raise ToolError(f"step {i}: link_text needs 'old_text' and 'new_text'")
        if s.get("part", "slide") not in PARTS:
            raise ToolError(f"step {i}: bad part {s.get('part')!r}")


def _short(res: dict) -> dict:
    keep = ("ok", "reverted", "failures", "error", "blockers", "saved", "skipped", "group_id", "new_id", "mapping",
            "slides", "pieces", "xml_check", "timeline_check", "warnings", "note", "lint", "message", "title_id",
            "text_boxes", "changed", "failed", "unfixable", "applied", "checker", "order_now", "deviation",
            "layout_change", "old_text", "new_text", "address", "before", "after", "warning",
            "hidden_text_marked_decorative", "variant_used", "unchanged", "needs_attention", "needs_distinct_title",
            "still_duplicate", "not_machine_checked", "table_alt_text_not_kept", "height_before", "height_after",
            "formatting_reset", "not_accepted", "issues_left", "hidden_text_failed", "extended_to_whole_run", "reads_now",
            "titles_not_marked")
    out = {k: res[k] for k in keep if k in res}
    if isinstance(res.get("visual_check"), dict):
        v = res["visual_check"]
        out["visual_changed_px"] = v.get("changed_pixels", {k: x.get("changed_pixels") for k, x in v.items()
                                                             if isinstance(x, dict)})
    if isinstance(out.get("xml_check"), dict):
        out["xml_check"] = out["xml_check"].get("status")
    if isinstance(res.get("changed"), list) and res["changed"] and isinstance(res["changed"][0], dict) \
            and "id" in res["changed"][0]:
        out["changed"] = [c.get("id") for c in res["changed"]]
    return out


def _leaves(s) -> set[int]:
    if s.Type != 6:
        return {s.Id}
    out = set()
    for i in range(1, s.GroupItems.Count + 1):
        out |= _leaves(s.GroupItems(i))
    return out


def _already_grouped(pres, slide: int, part: str, ids: list[int]) -> int | None:
    """The id of the top-level group whose pieces are exactly `ids` (a group step that already ran), else
    None. A group made of groups counts by its leaf pieces."""
    from .inventory import find_shape, get_part
    cont = get_part(pres, slide, part)
    if any(top_level_index(cont, i) is not None for i in ids):
        return None
    want = set()
    tops = set()
    for i in ids:
        s = safe(lambda i=i: find_shape(cont, i))
        if s is None:
            return None
        g = s
        while safe(lambda: g.ParentGroup) is not None:
            g = g.ParentGroup
        tops.add(g.Id)
        want |= _leaves(s) if s.Type == 6 else {s.Id}
    if len(tops) != 1:
        return None
    gid = tops.pop()
    top = cont.Shapes(top_level_index(cont, gid))
    return gid if _leaves(top) == want else None


def _simple_edit(pres, fn, *args) -> dict:
    """alt / decorative / clear_alt as a transaction too: a failure part-way leaves nothing behind."""
    session.checkpoint(pres)
    try:
        res = fn(pres, *args)
        session.save(pres)
        res["ok"] = True
        return res
    except Exception as exc:
        session.revert(pres)
        return {"ok": False, "reverted": True, "error": f"{type(exc).__name__}: {com_err(exc)}"}


def _wcag_step(pres, op: str, step: dict, m: dict) -> dict:
    from . import contrast, links, readingorder, tables, titles
    slide = int(step["slide"]) if "slide" in step else None
    if op in DECK_OPS and slide is not None and not step.get("slides"):
        step = {**step, "slides": [slide]}  # a deck-level op given one slide works on that slide only
    one = (lambda: _resolve(m, slide, [step["id"]])[0]) if "id" in step else None
    if op == "promote_title":
        if titles.title_text(pres.Slides(slide)):
            return {"ok": True, "skipped": "slide already has a title"}
        return titles.promote_to_title(pres, slide, one())
    if op == "hidden_title":
        if titles.title_text(pres.Slides(slide)):
            return {"ok": True, "skipped": "slide already has a title"}
        return titles.add_hidden_title(pres, slide, step["text"])
    if op == "set_title":
        if titles.title_text(pres.Slides(slide)) == step["text"]:
            return {"ok": True, "skipped": "title already set"}
        return titles.set_slide_title(pres, slide, step["text"])
    if op == "titles_unique":
        r = titles.make_titles_unique(pres, step.get("slides"), step.get("pattern", "{title} ({n} of {N})"))
        return {"ok": not r["failed"], **r}
    if op == "table_to_shapes":
        sl = pres.Slides(slide)
        if top_level_index(sl, one()) is None:
            tag = f"table{one()}"
            if any(safe(lambda i=i: sl.Shapes(i).Tags("A11Y_BLOCK")) == tag for i in range(1, sl.Shapes.Count + 1)):
                return {"ok": True, "skipped": "table already converted"}
            raise ToolError(f"no table #{one()} on slide {slide} (and no shapes converted from it)")
        return tables.table_to_shapes(pres, slide, one(), step.get("reading", "auto"))
    if op == "unmerge_table":
        return tables.unmerge_table(pres, slide, one(), bool(step.get("repeat_text", True)))
    if op == "table_header":
        return tables.set_table_header(pres, slide, one())
    if op == "link_text":
        holder = links.link_holder(pres.Slides(slide), one(), step.get("cell"), step["new_text"], step["old_text"])
        want = links._normalize(step["address"]) if step.get("address") else None
        if any(l["text"].strip() == step["new_text"].strip() and (want is None or l["address"] == want)
               for l in links._links_in(holder)):
            return {"ok": True, "skipped": "link text already replaced"}
        return links.set_link_text(pres, slide, one(), step["old_text"], step["new_text"], step.get("address"),
                                   step.get("cell"))
    if op == "contrast":
        r = contrast.fix_contrast(pres, step.get("slides"), step.get("strategy", "auto"))
        return {"ok": not r["failed"] and not r["unfixable"], **r}
    if op == "reading_order":
        return readingorder.set_reading_order(pres, slide, _resolve(m, slide, step["ids"]), strict=bool(step.get("strict")))
    if op == "fix_reading_order":
        orders = {int(k): _resolve(m, int(k), v) for k, v in (step.get("orders") or {}).items()} or None
        r = readingorder.fix_reading_order(pres, step.get("slides"), orders, bool(step.get("use_checker", True)),
                                           hidden_text=step.get("hidden_text", "decorative"))
        return {"ok": not r["failed"] and not r.get("hidden_text_failed"), **r}
    raise ToolError(f"unknown op {op}")


def run_step(deck_path: str, step: dict) -> dict:
    pres = session.get_pres(deck_path)
    m = load_map(pres)
    before = dict(m)
    op, part = step["op"], step.get("part", "slide")
    if op == "note":
        return {"ok": True, "skipped": "note"}
    if op in OPS - {"group", "alt", "decorative", "clear_alt", "flatten_group", "placeholder_to_picture",
                    "move_to_slides"}:
        return _wcag_step(pres, op, step, m)
    slide = int(step["slide"])
    if op == "group" and _already_grouped(pres, slide, part, _resolve(m, slide, step["ids"])):
        res = {"ok": True, "skipped": "already grouped",
               "group_id": _already_grouped(pres, slide, part, _resolve(m, slide, step["ids"]))}
    elif op == "group":
        res = ops.group_shapes(pres, slide, _resolve(m, slide, step["ids"]), step.get("alt_text"),
                               bool(step.get("decorative")), step.get("name"), part,
                               bool(step.get("allow_visual_change")))
    elif op == "alt":
        res = _simple_edit(pres, ops.set_alt_text, slide, _resolve(m, slide, step["ids"]), step["text"], part)
    elif op == "decorative":
        res = _simple_edit(pres, ops.mark_decorative, slide, _resolve(m, slide, step["ids"]), part)
    elif op == "clear_alt":
        res = _simple_edit(pres, ops.clear_alt_text, slide, _resolve(m, slide, step["ids"]), part)
    elif op == "flatten_group":
        gid = int(step["id"])
        if top_level_index(pres.Slides(slide), gid) is None and step.get("pieces") and all(
                top_level_index(pres.Slides(slide), p) for p in step["pieces"]):
            res = {"ok": True, "skipped": "already flattened"}
        else:
            res = structure.flatten_group(pres, slide, gid, bool(step.get("allow_alt_loss")))
    elif op == "placeholder_to_picture":
        old = int(step["id"])
        k = f"{slide}:slide:{old}"
        if k in m and top_level_index(pres.Slides(slide), old) is None and \
                top_level_index(pres.Slides(slide), int(m[k])) is not None:
            res = {"ok": True, "skipped": "already converted", "new_id": int(m[k])}
        else:
            res = structure.placeholder_to_picture(pres, slide, old)
            if res.get("ok"):
                _put(m, slide, "slide", old, res["new_id"])
    elif op == "move_to_slides":
        ids = [int(i) for i in step["ids"]]
        from .inventory import get_part
        src = get_part(pres, slide, part)
        if all(top_level_index(src, i) is None for i in ids) and \
                any(k.endswith(f":{part}:{ids[0]}") for k in m):
            res = {"ok": True, "skipped": "already moved"}
        else:
            res = structure.move_to_slides(pres, slide, ids, part)
            if res.get("ok"):
                for n, mp in res["mapping"].items():
                    for old, new in mp.items():
                        _put(m, int(n), part, int(old), int(new))
    else:
        raise ToolError(f"unknown op {op}")
    if m != before:
        _save_map(session.get_pres(deck_path), m)
    return res


def fill_from_sheet(steps: list[dict], sheet_path: str | None, column: str = "Alt Text") -> list[dict]:
    """Replace alt_text/text "ROW" with that spreadsheet row's alt text, verbatim."""
    if not any(s.get("alt_text") == "ROW" or s.get("text") == "ROW" for s in steps):
        return steps
    if not sheet_path:
        raise ToolError('steps use "ROW" alt text; pass sheet_path')
    from . import sheet
    rows = {r["row"]: r for r in sheet.read(sheet_path)["rows"]}
    out = []
    for s in steps:
        s = dict(s)
        for k in ("alt_text", "text"):
            if s.get(k) == "ROW":
                r = rows.get(int(s.get("row", -1)))
                if not r or not str(r.get(column, "")).strip():
                    raise ToolError(f"row {s.get('row')} has no '{column}' text")
                s[k] = str(r[column]).strip()
        out.append(s)
    return out


def apply_plan(deck_path: str, steps: list[dict], stop_on_failure: bool = True, sheet_path: str | None = None) -> dict:
    validate(steps)
    steps = fill_from_sheet(steps, sheet_path)
    results, ok = [], True
    for i, step in enumerate(steps, 1):
        p0 = session.get_pres(deck_path)
        if session.is_working_copy(p0.FullName) and not p0.Saved:
            session.checkpoint(p0)  # a failing step may revert: never let it take earlier, unsaved edits along
        try:
            res = run_step(deck_path, step)
        except Exception as exc:  # never lose the record of steps already applied
            res = {"ok": False, "error": f"{type(exc).__name__}: {com_err(exc)}"}
            pres = session.get_pres(deck_path)
            if not pres.Saved and session.is_working_copy(pres.FullName):
                session.revert(pres)
                res["reverted"] = True
        results.append({"step": i, "row": step.get("row"), "op": step["op"], "slide": step.get("slide"), **_short(res)})
        if not res.get("ok"):
            ok = False
            if stop_on_failure:
                break
    return {"ok": ok, "done": len(results), "of": len(steps), "results": results}


def select_rows(steps: list[dict], rows: list) -> list[dict]:
    want = {int(r) for r in rows}
    out = [s for s in steps if int(s.get("row", -1)) in want]
    missing = sorted(want - {int(s.get("row", -1)) for s in out})
    if missing:
        raise ToolError(f"no plan steps for rows {missing}")
    return out
