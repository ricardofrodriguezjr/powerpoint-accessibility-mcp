"""Click-by-click Slide Show comparison of every animated slide: working copy vs untouched original.

    .venv\\Scripts\\python.exe tests\\playback_all.py ORIGINAL.pptx WORKING.pptx OUT.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppta11y import playback, session  # noqa: E402
from ppta11y.com import COM, app, retry  # noqa: E402


def main(original: str, working: str, out: str) -> int:
    def animated_slides():
        p = app().Presentations.Open(str(Path(original).resolve()), -1, 0, 0)
        n = [i for i in range(1, p.Slides.Count + 1) if p.Slides(i).TimeLine.MainSequence.Count]
        retry(p.Close)
        return n

    slides = COM.run(animated_slides)
    print("animated slides:", slides, flush=True)
    results = {}
    for n in slides:
        def one(n=n):
            p = session.get_pres(working)
            return playback.compare_playback(p, n, original)
        try:
            r = COM.run(one)
            results[n] = {"ok": r["ok"], "clicks": r["clicks"], "changed_px": [f["changed_px"] for f in r["frames"]],
                          "image": r["image"]}
        except Exception as exc:
            results[n] = {"ok": False, "error": str(exc)}
        print(n, results[n].get("ok"), results[n].get("clicks"), results[n].get("changed_px"), flush=True)
        Path(out).write_text(json.dumps(results, indent=1), encoding="utf8")
    bad = [n for n, r in results.items() if not r.get("ok")]
    print(f"{len(results)} animated slides played; differing: {bad}")
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
