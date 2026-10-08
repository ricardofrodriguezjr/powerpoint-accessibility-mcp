"""Read PowerPoint's own Accessibility Assistant (Review > Check Accessibility) through UI Automation.

There is no COM API for the checker's results. The pane is opened with ExecuteMso("AccessibilityChecker"),
its category buttons ("Missing alt text - 12") are read with UI Automation, and individual issues are
itemised by stepping "Next issue" and reading which shape PowerPoint selects.
"""
from __future__ import annotations

import re
import time
import warnings

from .com import ToolError, app, safe

PANE = "Accessibility Assistant"
CAT_RX = re.compile(r"^(?P<name>.+?) - (?P<n>\d+)$")
HDR_RX = re.compile(r"^(?P<k>\d+) of (?P<n>\d+)")


def _desktop():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        from pywinauto import Desktop  # imported lazily: slow, and only needed here
    return Desktop(backend="uia")


def _main_window(pres):
    """The top-level window whose title is exactly '<document window caption> - PowerPoint' (the caption
    already reflects hidden file extensions, read-only markers, etc.)."""
    cap = safe(lambda: pres.Windows(1).Caption, pres.Name)
    hits = [w for w in _desktop().windows()
            if w.window_text() == f"{cap} - PowerPoint" or w.window_text().startswith(f"{cap} - ")]
    hits = [w for w in hits if not w.window_text().startswith("PowerPoint Slide Show")]
    if len(hits) != 1:
        raise ToolError(f"found {len(hits)} PowerPoint windows titled '{cap} - ...'; close duplicates and retry")
    return hits[0]


def _pane(main):
    for c in main.descendants(control_type="Custom"):
        if c.element_info.name == PANE:
            return c
    return None


def _by_aid(pane, aid):
    for c in pane.descendants():
        if c.element_info.automation_id == aid:
            return c
    return None


def _open_pane(pres):
    win = pres.Windows(1)
    safe(lambda: setattr(win, "WindowState", 1))  # ppWindowNormal (minimised windows have no UI tree)
    win.Activate()
    main = _main_window(pres)
    pane = _pane(main)
    if pane is None:
        app().CommandBars.ExecuteMso("AccessibilityChecker")
        for _ in range(40):
            time.sleep(0.25)
            pane = _pane(main)
            if pane is not None:
                break
    if pane is None:
        raise ToolError("the Accessibility pane did not open. This reader supports the Microsoft 365 "
                        "'Accessibility Assistant'; on older builds run Review > Check Accessibility by hand.")
    back = _by_aid(pane, "ASSISTANT_DRILLIN_BACK")
    if back is not None:
        back.invoke()
        time.sleep(0.8)
        pane = _pane(main)
    return main, pane


def _categories(pane) -> dict:
    cats = {}
    for b in pane.descendants(control_type="Button"):
        m = CAT_RX.match(b.element_info.name or "")
        if m:
            cats[m["name"]] = int(m["n"])
    return cats


def _selected(pres):
    w = pres.Windows(1)
    out = {"slide": safe(lambda: w.View.Slide.SlideIndex)}
    sel = safe(lambda: w.Selection)
    if sel is not None and safe(lambda: sel.Type, 0) >= 2:
        sr = sel.ShapeRange
        shapes = []
        for i in range(1, sr.Count + 1):
            s = sr(i)
            shapes.append({"id": s.Id, "name": s.Name, "type": s.Type,
                           **({"in_group": s.ParentGroup.Id} if safe(lambda: s.ParentGroup.Id) else {})})
        out["shapes"] = shapes
        if safe(lambda: sel.HasChildShapeRange):
            cr = sel.ChildShapeRange
            out["child_shapes"] = [{"id": cr(i).Id, "name": cr(i).Name} for i in range(1, cr.Count + 1)]
    return out


def run(pres, itemize: list[str] | None = None, max_items: int = 400, step_delay: float = 0.6) -> dict:
    """Open the Accessibility Assistant and return its category counts; for each category named in
    `itemize`, step through every issue and record the slide and shape PowerPoint selects."""
    main, pane = _open_pane(pres)
    time.sleep(1.5)  # the assistant re-scans after edits
    pane = _pane(main)
    cats = _categories(pane)
    quick = [t.element_info.name for t in pane.descendants(control_type="Text")
             if "can be marked as decorative" in (t.element_info.name or "")]
    result = {"deck": pres.FullName, "categories": cats,
              "issues_total": sum(cats.values()), **({"quick_fix": quick[0]} if quick else {}), "items": {}}
    for cat in itemize or []:
        n = cats.get(cat)
        if not n:
            result["items"][cat] = []
            continue
        btn = next((b for b in pane.descendants(control_type="Button")
                    if (b.element_info.name or "").startswith(cat + " - ")), None)
        if btn is None:
            continue
        btn.invoke()
        time.sleep(1.0)
        items, seen = [], set()
        try:
            for k in range(min(n, max_items)):
                pane = _pane(main)
                if pane is None:
                    break
                hdr = next((t.element_info.name for t in pane.descendants(control_type="Text")
                            if HDR_RX.match(t.element_info.name or "")), "")
                m = HDR_RX.match(hdr)
                if m and (int(m["k"]) in seen or int(m["n"]) != n):
                    result.setdefault("notes", []).append(
                        f"{cat}: list changed while stepping (header '{hdr}'); stopped after {len(items)} items")
                    break
                if m:
                    seen.add(int(m["k"]))
                items.append({"n": k + 1, "header": hdr, **_selected(pres)})
                if k < min(n, max_items) - 1:
                    nxt = _by_aid(pane, "ASSISTANT_DRILLIN_NEXT")
                    if nxt is None:
                        break
                    nxt.invoke()
                    time.sleep(step_delay)
        finally:  # always leave drill-in mode
            pane = _pane(main)
            back = _by_aid(pane, "ASSISTANT_DRILLIN_BACK") if pane is not None else None
            if back is not None:
                back.invoke()
                time.sleep(0.8)
        result["items"][cat] = items
        if n > max_items:
            result["items"][cat + " (truncated)"] = f"stopped after {max_items} of {n}"
        pane = _pane(main)
    return result
