"""Whole-deck visual regression: export every slide of two decks with PowerPoint and pixel-diff them.

    .venv\\Scripts\\python.exe tests\\deck_diff.py ORIGINAL.pptx WORKING.pptx OUTDIR
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppta11y import render  # noqa: E402
from ppta11y.com import COM, app, retry  # noqa: E402
from ppta11y.session import find_open  # noqa: E402


def main(a: str, b: str, outdir: str, tolerance_px: int = 20) -> int:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)

    def export_all(path, tag):
        p = find_open(path)
        opened = p is None
        if opened:
            p = app().Presentations.Open(str(Path(path).resolve()), -1, 0, 0)
        files = [render.export(p, n, out / f"{tag}_{n:02d}.png", 1280) for n in range(1, p.Slides.Count + 1)]
        if opened:
            retry(p.Close)
        return files

    A = COM.run(export_all, a, "orig")
    B = COM.run(export_all, b, "work")
    if len(A) != len(B):
        print(f"slide count differs: {len(A)} vs {len(B)}")
        return 1
    res = []
    for n, (x, y) in enumerate(zip(A, B), 1):
        d = render.diff(x, y)
        res.append({"slide": n, "changed_px": d["changed_pixels"], "box": d["changed_box_px"]})
    bad = [r for r in res if r["changed_px"] > tolerance_px]
    (out / "deck_diff.json").write_text(json.dumps(res, indent=1), encoding="utf8")
    print(f"{len(res)} slides compared; {len(bad)} differ by more than {tolerance_px} px")
    for r in bad:
        print("  DIFF", r)
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
