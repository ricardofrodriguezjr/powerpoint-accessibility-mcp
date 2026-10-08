"""Deck-wide alt-text audit (our own rules, aligned with PowerPoint's Accessibility Checker)."""
from __future__ import annotations

import csv
import json
from collections import Counter, defaultdict

from .com import safe
from .inventory import (iter_shapes, kind, layout_key, needs_alt, slide_title, slides_showing_master,
                        slides_using_layout, text_of, timeline_targets)
from .ops import is_junk, lint_alt
from .session import out_dir, stamp

STATUS_HELP = {
    "missing": "needs alt text (or Mark as decorative); PowerPoint's checker flags it",
    "review": "shape without text: give it alt text if it carries meaning, otherwise mark decorative",
    "junk": "has alt text a screen reader would read out, but it is a file name/URL/auto-text",
    "ok": "has alt text",
    "decorative": "marked decorative",
    "text": "text is read directly; no alt text needed",
    "table": "native table; read cell by cell (check it has a header row)",
}


def classify(s) -> dict:
    alt = (s.AlternativeText or "").strip()
    dec = bool(safe(lambda: s.Decorative, 0))
    need = needs_alt(s)
    if s.Type == 19 or safe(lambda: s.HasTable) == -1:
        header = safe(lambda: bool(s.Table.FirstRow))
        return {"status": "table", "header_row": header}
    if dec:
        out = {"status": "decorative"}
        if alt:
            out["warning"] = "decorative but still has alt text (clear it)"
        return out
    if alt:
        if is_junk(alt):
            return {"status": "junk", "lint": lint_alt(alt)}
        lint = lint_alt(alt)
        return {"status": "ok", **({"lint": lint} if lint else {})}
    if need == "required":
        return {"status": "missing"}
    if need == "review":
        return {"status": "review"}
    return {"status": "text"}


def _row(where: str, s, extra: dict) -> dict:
    t = text_of(s).strip()
    return {"where": where, "id": s.Id, "name": s.Name, "kind": kind(s),
            "alt": (s.AlternativeText or "")[:200], **({"text": t[:80]} if t else {}), **extra}


def _needs_work(r: dict) -> bool:
    return (r["status"] in ("missing", "review", "junk") or "warning" in r or bool(r.get("junk_alt_inside_group"))
            or (r["status"] == "table" and r.get("header_row") is False))


def audit(pres, include_inherited: bool = True, include_ok: bool = False) -> dict:
    rows, titles = [], []
    for n in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(n)
        targets, _ = timeline_targets(sl)
        if not slide_title(sl):
            titles.append(n)
        for i in range(1, sl.Shapes.Count + 1):
            s = sl.Shapes(i)
            extra = {"slide": n, **classify(s)}
            if targets.get(s.Id):
                extra["animated"] = targets[s.Id]
            if s.Type == 6:
                members = [m for m, _, _ in iter_shapes(_GroupShapes(s))]
                inner = [m for m in members if is_junk(m.AlternativeText)]
                if inner:
                    extra["junk_alt_inside_group"] = [{"id": m.Id, "alt": m.AlternativeText[:80]} for m in inner]
                anim_inside = sum(targets.get(m.Id, 0) for m in members)
                if anim_inside:
                    extra["animated_members"] = anim_inside
            row = _row(f"slide {n}", s, extra)
            if include_ok or _needs_work(row) or "lint" in row:
                rows.append(row)

    inherited = []
    if include_inherited:
        seen = set()
        for n in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(n)
            lay = sl.CustomLayout
            key = ("layout",) + layout_key(lay)
            if key not in seen:
                seen.add(key)
                shown = slides_using_layout(pres, lay, shown_only=True)
                for i in range(1, lay.Shapes.Count + 1):
                    s = lay.Shapes(i)
                    if s.Type == 14:
                        continue
                    r = _row(f"layout '{lay.Name}'", s, {"slide": shown[0] if shown else n, "part": "layout",
                                                          "shown_on": shown, **classify(s)})
                    if not shown:
                        r["note"] = "not shown on any slide (all its slides hide background graphics)"
                    if include_ok or (_needs_work(r) and shown):
                        inherited.append(r)
            m = sl.Master
            key = ("master", sl.Design.Index)
            if key not in seen:
                seen.add(key)
                shown = slides_showing_master(pres, sl.Design.Index)
                for i in range(1, m.Shapes.Count + 1):
                    s = m.Shapes(i)
                    if s.Type == 14:
                        continue
                    r = _row(f"master '{m.Name}'", s, {"slide": shown[0] if shown else n, "part": "master",
                                                        "shown_on": shown, **classify(s)})
                    if include_ok or (_needs_work(r) and shown):
                        inherited.append(r)

    work = [r for r in rows if _needs_work(r)]
    counts = Counter(r["status"] for r in rows)
    per_slide = defaultdict(list)
    for r in work:
        per_slide[r["slide"]].append(r["id"])
    return {
        "deck": pres.FullName,
        "slides": pres.Slides.Count,
        "summary": {
            "slides_objects_needing_work": len(work),
            "missing_review_junk": sum(counts[k] for k in ("missing", "review", "junk")),
            "by_status": dict(counts),
            "other_warnings": len([r for r in work if r["status"] not in ("missing", "review", "junk")]),
            "layout_master_objects_needing_work": len([r for r in inherited if _needs_work(r)]),
            "slides_missing_title": titles,
        },
        "status_help": STATUS_HELP,
        "todo_by_slide": dict(per_slide),
        "objects": rows,
        "inherited_objects": inherited,
    }


class _GroupShapes:
    """Adapter so iter_shapes can walk a group's members like a Shapes collection."""

    def __init__(self, g):
        self._items = g.GroupItems
        self.Count = self._items.Count

    def __call__(self, i):
        return self._items(i)


def write_report(pres, report: dict) -> dict:
    d = out_dir(pres)
    base = d / f"audit_{stamp()}"
    (base.with_suffix(".json")).write_text(json.dumps(report, indent=1, default=str), encoding="utf8")
    cols = ["where", "slide", "id", "name", "kind", "status", "alt", "text", "animated", "lint", "warning",
            "junk_alt_inside_group", "header_row", "shown_on", "note"]
    with open(base.with_suffix(".csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in report["objects"] + report["inherited_objects"]:
            w.writerow({k: (json.dumps(v) if isinstance(v, (list, dict)) else v) for k, v in r.items()})
    return {"json": str(base.with_suffix(".json")), "csv": str(base.with_suffix(".csv"))}
