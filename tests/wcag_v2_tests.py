"""End-to-end test of the v2 WCAG tools on a synthetic deck built through COM.

    .venv\\Scripts\\python.exe tests\\wcag_v2_tests.py <scratch folder>      -> prints RESULT: PASS / FAIL

Builds a deck with: a Blank-layout slide whose heading is a text box, a 4x4 table with a body merge, light
grey text on white, a raw URL link alone in its paragraph; a slide with two overlapping shapes (reading order
must not restack them visibly) plus a footer-like text box; two slides with the same title. Then runs the
tools and checks every claim (titles, uniqueness, tables, links, contrast, reading order, invisibility).
"""
from __future__ import annotations

import json
import os
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppta11y import contrast, links, readingorder, render, session, tables, titles, wcag  # noqa: E402
from ppta11y.com import COM, app  # noqa: E402

TRUE = -1
LAMBDA_TITLE = "Choosing 𝜆 by CV"  # a math-italic letter: 2 UTF-16 units in PowerPoint


def pure_checks(ok) -> None:
    """Checks that need no PowerPoint."""
    import xml.etree.ElementTree as ET
    from PIL import Image, ImageDraw
    # dead links: per paragraph, line breaks kept, table cells found
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    dead = f'<a:r xmlns:a="{A}" xmlns:r="{R}"><a:rPr><a:hlinkClick r:id=""/></a:rPr><a:t>{{}}</a:t></a:r>'
    par = ET.fromstring(f'<a:p xmlns:a="{A}">' + dead.format("www.example.edu/") + f'<a:br xmlns:a="{A}"/>'
                       + dead.format("stats") + '</a:p>')
    ok("dead link keeps its line break", links._dead_runs(par) == ["www.example.edu/\x0bstats"], links._dead_runs(par))
    # a data table with a full-height merged column and a spanning 'Price' header is not a tree
    regs = []
    for r in range(1, 6):
        regs.append({"r0": r, "r1": r, "c0": 1, "c1": 1, "text": ["Item", "", "Bolt", "Nut", "Washer"][r - 1]})
        regs.append({"r0": r, "r1": r, "c0": 2, "c1": 3, "text": "Description" if r == 1 else "x"})
        if r == 1:
            regs.append({"r0": 1, "r1": 1, "c0": 4, "c1": 5, "text": "Price"})
        else:
            regs += [{"r0": r, "r1": r, "c0": 4, "c1": 4, "text": "2023" if r == 2 else "1.10"},
                     {"r0": r, "r1": r, "c0": 5, "c1": 5, "text": "2024" if r == 2 else "0.40"}]
    ok("data table with merged column is not a hierarchy", tables._tree(regs) is None)
    tree = [{"r0": 1, "r1": 1, "c0": 1, "c1": 4, "text": "Course"},
            {"r0": 2, "r1": 2, "c0": 1, "c1": 2, "text": "Part A"}, {"r0": 2, "r1": 2, "c0": 3, "c1": 4, "text": "Part B"},
            {"r0": 3, "r1": 3, "c0": 1, "c1": 1, "text": "Intro"}, {"r0": 3, "r1": 3, "c0": 2, "c1": 2, "text": "Models"},
            {"r0": 3, "r1": 3, "c0": 3, "c1": 3, "text": "Trees"}, {"r0": 3, "r1": 3, "c0": 4, "c1": 4, "text": "Nets"}]
    ok("roadmap drawn as a table is a hierarchy", tables._tree(tree) is not None)
    # low-contrast text: the glyphs are not taken for the background
    img = Image.new("RGB", (200, 60), (255, 255, 255))
    d = ImageDraw.Draw(img)
    for x in range(10, 190, 6):
        d.rectangle([x, 15, x + 3, 45], fill=(0xE6, 0xE6, 0xE6))
    bg = contrast._sample_bg(img, (0, 0, 200, 60), 1, (0xE6, 0xE6, 0xE6))
    ok("light-grey text samples a white background", bg and min(min(b) for b in bg) > 250, bg)


def build(path: Path) -> None:
    a = app()
    p = a.Presentations.Add(TRUE)
    try:
        # 1: Blank layout, heading text box, table with a merge, low contrast text, raw URL
        s1 = p.Slides.Add(1, 12)
        h = s1.Shapes.AddTextbox(1, 40, 20, 600, 40)
        h.TextFrame.TextRange.Text = "Course roadmap"
        h.TextFrame.TextRange.Font.Size = 28
        h.TextFrame.TextRange.Font.Bold = TRUE
        t = s1.Shapes.AddTable(4, 4, 40, 80, 400, 160)
        for r in range(1, 5):
            for c in range(1, 5):
                t.Table.Cell(r, c).Shape.TextFrame.TextRange.Text = f"r{r}c{c}"
        t.Table.Cell(2, 2).Merge(t.Table.Cell(2, 3))
        g = s1.Shapes.AddTextbox(1, 480, 90, 300, 40)
        g.TextFrame.TextRange.Text = "Light grey note"
        g.TextFrame.TextRange.Font.Size = 14
        g.TextFrame.TextRange.Font.Color.RGB = 0xBBBBBB
        u = s1.Shapes.AddTextbox(1, 480, 160, 400, 60)
        url = "https://example.com/very/long/path/to/a/page?ref=abc&u=123"
        u.TextFrame.TextRange.Text = "Read more:\r" + url
        u.TextFrame.TextRange.Paragraphs(2).ActionSettings(1).Hyperlink.Address = url
        # 2: overlapping shapes + a footer-like box at the bottom, read in a bad order
        s2 = p.Slides.Add(2, 2)
        s2.Shapes(1).TextFrame.TextRange.Text = LAMBDA_TITLE
        s2.Shapes(2).TextFrame.TextRange.Text = "Body text"
        a1 = s2.Shapes.AddShape(1, 500, 200, 150, 100)
        a1.Fill.ForeColor.RGB = 0x0000FF
        a1.AlternativeText = "red box"
        b1 = s2.Shapes.AddShape(1, 560, 240, 150, 100)
        b1.Fill.ForeColor.RGB = 0xFF0000
        b1.AlternativeText = "blue box overlapping the red one"
        s2.Shapes(1).ZOrder(0)  # title to the front: read last (bad order)
        # 3: duplicate title
        s3 = p.Slides.Add(3, 2)
        s3.Shapes(1).TextFrame.TextRange.Text = LAMBDA_TITLE
        s3.Shapes(2).TextFrame.TextRange.Text = "More body text"
        # a vertical axis label (rotated 270 degrees) in dark text, next to a dark bar: measured on white
        lab = s3.Shapes.AddTextbox(1, 560, 300, 200, 30)
        lab.TextFrame.TextRange.Text = "Test error rate"
        lab.TextFrame.TextRange.Font.Size = 14
        lab.TextFrame.TextRange.Font.Color.RGB = 0x333333
        lab.TextFrame.WordWrap = 0
        lab.Rotation = 270
        bar = s3.Shapes.AddShape(1, 700, 300, 40, 200)
        bar.Fill.ForeColor.RGB = 0x41208A
        bar.Line.Visible = 0
        bar.AlternativeText = "Bar"
        # 4: Title Only layout, title left empty, the heading drawn as a text box over it; an answer under a
        #    navy cover that leaves on click (exit fade)
        s4 = p.Slides.Add(4, 11)
        tp = s4.Shapes(1)
        hb = s4.Shapes.AddTextbox(1, tp.Left, tp.Top, tp.Width, tp.Height)
        hb.TextFrame.TextRange.Text = "Quiz question"
        hb.TextFrame.TextRange.Font.Size = 32
        ans = s4.Shapes.AddTextbox(1, 100, 250, 300, 50)
        ans.TextFrame.TextRange.Text = "Answer: 42"
        ans.TextFrame.TextRange.Font.Size = 20
        ans.TextFrame.TextRange.Font.Color.RGB = 0
        cov = s4.Shapes.AddShape(1, 80, 230, 360, 100)
        cov.Fill.ForeColor.RGB = 0x502010
        cov.Line.Visible = 0
        cov.AlternativeText = "Cover hiding the answer"
        e = s4.TimeLine.MainSequence.AddEffect(cov, 10, 0, 1)  # fade, on click
        e.Exit = TRUE
        p.SaveAs(str(path))
    finally:
        p.Close()


def main(folder: str) -> int:
    folder = Path(folder).resolve()
    folder.mkdir(parents=True, exist_ok=True)
    src = folder / "v2_src.pptx"
    wc = folder / "v2_src_accessible.pptx"
    for f in (src, wc):
        if f.exists():
            f.unlink()
    checks = []

    def ok(name, cond, info=""):
        checks.append((name, bool(cond), info))
        print(("PASS " if cond else "FAIL ") + name + (f"  {info}" if info else ""))

    pure_checks(ok)

    def run():
        build(src)
        session.open_working_copy(str(src), fresh=True)
        p = session.find_open(str(wc))
        # titles
        rep = titles.title_report(p)
        cand = next(m for m in rep["missing"] if m["slide"] == 1)["candidate"]
        ok("title candidate is the heading", cand and cand["text"] == "Course roadmap", cand)
        r = titles.promote_to_title(p, 1, cand["id"])
        ok("promote_to_title ok and invisible", r.get("ok") and r["visual_check"]["changed_pixels"] <= 20,
           {k: r.get(k) for k in ("ok", "failures", "refused_settings")})
        p = session.find_open(str(wc))
        ok("slide 1 has title", titles.title_text(p.Slides(1)) == "Course roadmap")
        s4 = next(m for m in titles.title_report(p)["missing"] if m["slide"] == 4)
        ok("empty title placeholder reported as empty", s4["title_state"] == "empty", s4)
        try:
            titles.set_slide_title(p, 4, "Quiz question")
            refused = False
        except Exception as exc:  # noqa: BLE001
            refused = "promote_to_title" in str(exc)
        ok("set_slide_title refuses an empty placeholder over the heading", refused)
        r = titles.promote_to_title(p, 4, s4["candidate"]["id"])
        p = session.find_open(str(wc))
        ok("slide 4 heading promoted", r.get("ok") and titles.title_text(p.Slides(4)) == "Quiz question", r.get("failures"))
        u = titles.make_titles_unique(p, slides=[3])
        p = session.find_open(str(wc))
        ok("titles unique (whole run numbered, astral letter kept)",
           not titles.title_report(p)["duplicates"] and u.get("extended_to_whole_run") == [[2, 3]]
           and titles.title_text(p.Slides(2)) == LAMBDA_TITLE + " (1 of 2)", u)
        # tables
        tr = tables.table_report(p)
        tid = tr["tables"][0]["id"]
        ok("merged table found", tr["tables"][0]["merged_regions"] == 1, tr["tables"][0])
        r = tables.unmerge_table(p, 1, tid)
        p = session.find_open(str(wc))
        ok("unmerge ok", r.get("ok"), r.get("failures"))
        ok("no merged tables left", not tables.table_report(p)["summary"]["with_merged_cells"])
        # links
        lr = links.link_report(p)
        bad = [l for l in lr["links"] if l["unclear"]]
        ok("unclear link found", len(bad) == 1, lr["summary"])
        r = links.set_link_text(p, 1, bad[0]["shape_id"], bad[0]["text"], "Example page about paths")
        p = session.find_open(str(wc))
        ok("link text replaced, link kept", r.get("ok") and not links.link_report(p)["summary"]["unclear"], r)
        # contrast
        cr = contrast.contrast_report(p)
        ok("low contrast found", cr["failing"] >= 1, [f["text"] for f in cr["failures"]])
        rot = [x for x in contrast.slide_contrast(p, 3) if x.get("text") == "Test error rate"]
        ok("rotated label measured where it is drawn", rot and rot[0]["passes"] and rot[0]["bg_all"] == ["#FFFFFF"], rot)
        ansr = [x for x in contrast.slide_contrast(p, 4) if x.get("text") == "Answer: 42"]
        ok("text under an animated cover is measured with the cover gone", ansr and ansr[0]["passes"]
           and ansr[0].get("measured_states") and ansr[0]["bg_rgb"] == "#FFFFFF", ansr)
        r = contrast.fix_contrast(p)
        p = session.find_open(str(wc))
        ok("contrast fixed", not r["failed"] and contrast.contrast_report(p)["failing"] == 0, r)
        # reading order: overlapping boxes must keep their stacking
        items = readingorder.slide_items(p, 2)
        ids = {it["text"] or it["alt"]: it["id"] for it in items}
        title_id = next(it["id"] for it in items if it["role"] == "title")
        body_id = next(it["id"] for it in items if it["role"] == "placeholder")
        red, blue = ids["red box"], ids["blue box overlapping the red one"]
        r = readingorder.set_reading_order(p, 2, [title_id, body_id, blue, red])
        p = session.find_open(str(wc))
        now = [it["id"] for it in readingorder.slide_items(p, 2) if it["read"]]
        ok("restack ok, overlapping pair kept, invisible",
           r.get("ok") and now.index(red) < now.index(blue) and r["visual_check"]["changed_pixels"] <= 20
           and now[0] == title_id, {"now": now, "plan": r.get("plan"), "dev": r.get("deviation")})
        ok("a closest-invisible order is not accepted without a look", r.get("not_accepted") and
           readingorder.load_accepted(p).get(str(p.Slides(2).SlideID), {}).get("order") != now, r.get("not_accepted"))
        r = readingorder.fix_reading_order(p, use_checker=False)
        p = session.find_open(str(wc))
        ok("answer under an animated cover: closest order flagged for a look", 4 in r["needs_attention"],
           r["needs_attention"])
        for n in r["needs_attention"]:  # the agent checks reads_now, then accepts exactly that order
            acc = readingorder.set_reading_order(p, n, r["applied"][n]["order_now"])
            p = session.find_open(str(wc))
            ok(f"slide {n}: closest order accepted explicitly", acc.get("ok") and acc.get("unchanged"), acc)
        rep = readingorder.deck_report(p, write=False)
        ok("every slide matches its order", not rep["summary"]["slides_to_fix"], rep["summary"])
        a = wcag.wcag_audit(p, run_checker=True)
        ok("wcag_audit own rules clean", all(not v for v in a["remaining"].values() if v is not None), a["remaining"])
        ok("PowerPoint checker clean (except reading-order heuristic)",
           not {k for k in a["powerpoint_checker_open"] if k != "Check reading order"}, a["powerpoint_checker"])
        session.close(p, save_first=True)

    COM.run(run)
    failed = [c for c in checks if not c[1]]
    print(json.dumps({"checks": len(checks), "failed": [c[0] for c in failed]}, indent=1))
    print("RESULT: " + ("PASS" if not failed else "FAIL"))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else os.getcwd()))
