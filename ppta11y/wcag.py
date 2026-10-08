"""One deck-wide WCAG audit and the definition of "done".

Covers these WCAG 2.1 AA / Section 508 criteria (not all of AA; see NOT_CHECKED):
  1.1.1 alt text or decorative on every non-text object (audit.audit)
  1.3.1 / 2.4.6 every slide has a real title; titles unique; tables have a header row and no merged cells
  1.3.2 / 2.4.3 reading order: title first, content in visual order, footers last (or an accepted order)
  1.4.3 text contrast 4.5:1 (3:1 large text) against the rendered background
  2.4.4 link text says where the link goes; links have a target
PowerPoint's own Accessibility Assistant is run on top as a second opinion (its reading-order rule is a
heuristic and may not clear even when the order is right; that is reported, not hidden).
"""
from __future__ import annotations

import json

from . import audit as alt_audit
from . import contrast, links, readingorder, tables, titles
from .session import out_dir, stamp

# What this audit does not check: reported with every audit so nobody mistakes done=true for a certificate.
NOT_CHECKED = [
    "1.2 captions and transcripts for video or audio",
    "1.4.1 meaning shown only by colour",
    "1.4.11 contrast of chart lines, arrows and icons",
    "3.1.1 / 3.1.2 the language of the deck and of passages in another language",
    "2.2.2 slides or media that advance, move or loop on their own",
    "2.3.1 flashing content",
    "whether each alt text is accurate (a person should review it)",
]


def wcag_audit(pres, compare_to: str | None = None, run_checker: bool = True, check_contrast: bool = True) -> dict:
    alt = alt_audit.audit(pres, include_inherited=True)
    ti = titles.title_report(pres)
    ro = readingorder.deck_report(pres, compare_to=compare_to, write=True)
    tb = tables.table_report(pres)
    lk = links.link_report(pres)
    ct = contrast.contrast_report(pres) if check_contrast else None
    alt_work = [r for r in alt["objects"] + alt["inherited_objects"]
                if alt_audit._needs_work(r) and r["status"] != "table"]  # header rows are counted below
    remaining = {
        "alt_text (1.1.1)": len(alt_work),
        "slides without a title (2.4.2/2.4.6)": len(ti["missing"]),
        "slides with a duplicate title (2.4.6)": sum(len(d["slides"]) for d in ti["duplicates"]),
        "slides whose reading order should change (1.3.2)": len(ro["summary"]["slides_to_fix"]),
        "slides whose title is not read first (1.3.2)": len(ro["summary"]["title_not_first"]),
        "slides whose footer is read before content (1.3.2)": len(ro["summary"]["footer_before_content"]),
        "text runs below the contrast minimum (1.4.3)": ct["failing"] if ct else None,
        "tables with merged cells (1.3.1)": len(tb["summary"]["with_merged_cells"]),
        "tables without a header row (1.3.1)": len(tb["summary"]["without_header_row"]),
        "tables with two header rows (1.3.1)": len(tb["summary"]["multi_row_header"]),
        "unclear or dead links (2.4.4)": lk["summary"]["unclear"],
    }
    where = {
        "alt_text": {n: [r["id"] for r in alt_work if r.get("slide") == n and r.get("part", "slide") == "slide"]
                     for n in sorted({r["slide"] for r in alt_work if r.get("part", "slide") == "slide"})},
        "inherited_needing_work": [r for r in alt["inherited_objects"] if alt_audit._needs_work(r)],
        "titles_missing": ti["missing"], "titles_duplicate": ti["duplicates"],
        "reading_order": {k: v for k, v in ro["summary"].items() if k != "slides_checked"},
        "contrast": ct["failures"] if ct else None,
        "tables": tb["tables"], "links": [l for l in lk["links"] if l["unclear"]],
        "link_advice": [l for l in lk["links"] if l.get("advice")],
    }
    checker, skipped = None, [name for name, on in (("contrast", check_contrast), ("PowerPoint checker", run_checker))
                              if not on]
    if run_checker:
        from . import checker as chk
        try:
            checker = chk.run(pres)["categories"]
        except Exception as exc:  # noqa: BLE001 - the checker drives PowerPoint's UI; keep our own results
            from .com import com_err
            skipped.append(f"PowerPoint checker (failed: {type(exc).__name__}: {com_err(exc)[:120]}; run "
                           "run_accessibility_checker again)")
    own_ok = all(not v for v in remaining.values() if v is not None)
    checker_left = {k: v for k, v in (checker or {}).items() if v}
    caveats = []
    if ct and ct.get("not_machine_checked"):
        caveats.append({"chart text not measured for contrast (check by eye)": ct["not_machine_checked"]})
    out = {"deck": pres.FullName, "done": (own_ok and not checker_left) if not skipped else None,
           "skipped": skipped, "remaining": remaining, "caveats": caveats,
           "powerpoint_checker": checker, "powerpoint_checker_open": checker_left, "details": where,
           "reading_order_report": ro.get("report_files"), "not_checked": NOT_CHECKED}
    if own_ok and set(checker_left) == {"Check reading order"}:
        out["note"] = ("every rule is met; PowerPoint's 'Check reading order' is a heuristic prompt that can stay on "
                       "after a correct fix - review those slides with reading_order_report / reading_order_snapshot")
    p = out_dir(pres) / f"wcag_audit_{stamp()}.json"
    p.write_text(json.dumps(out, indent=1, ensure_ascii=False, default=str), encoding="utf8")
    out["report_file"] = str(p)
    return out
