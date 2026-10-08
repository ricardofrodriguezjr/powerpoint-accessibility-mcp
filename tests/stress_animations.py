"""Stress test: build a slide with many kinds of animation, group the animated shapes with
ops.group_shapes, and require that PowerPoint's saved animation XML is unchanged.

Run:  .venv\\Scripts\\python.exe tests\\stress_animations.py [outdir]
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ppta11y import animations as anim  # noqa: E402
from ppta11y import ops, session  # noqa: E402
from ppta11y.com import COM, TRUE, app, retry  # noqa: E402

FLY, WIPE, FADE, SPIN, PATH_DOWN, GROW, FILL_COLOR, APPEAR, ZOOM = 2, 22, 10, 61, 127, 59, 54, 1, 23
ON_CLICK, WITH_PREV, AFTER_PREV, ON_SHAPE = 1, 2, 3, 4


def build(path: str) -> dict:
    a = app()
    pres = a.Presentations.Add(TRUE)
    sl = pres.Slides.Add(1, 12)  # ppLayoutBlank
    S = sl.Shapes
    shapes = {}

    def box(key, t, x, y, w=90, h=60, text=None):
        s = S.AddShape(t, x, y, w, h)
        s.Name = key
        if text:
            s.TextFrame.TextRange.Text = text
        shapes[key] = s
        return s

    A = box("A_fly", 1, 40, 40)
    B = box("B_wipe", 9, 150, 40)
    C = S.AddTextbox(1, 260, 30, 200, 90)
    C.Name = "C_paragraphs"
    C.TextFrame.TextRange.Text = "First point\rSecond point\rThird point"
    shapes["C_paragraphs"] = C
    D = box("D_spin", 92, 480, 40)
    E = box("E_path", 1, 40, 150)
    F = box("F_exit", 5, 150, 150)
    G = box("G_grow", 1, 260, 150)
    H = box("H_color", 1, 370, 150)
    I = box("I_interactive", 1, 480, 150)
    J = box("J_dim", 1, 40, 260, text="dim me")
    K = S.AddTextbox(1, 150, 260, 200, 40)
    K.Name = "K_letters"
    K.TextFrame.TextRange.Text = "By letter"
    shapes["K_letters"] = K
    N = box("N_outside", 1, 600, 300)            # animated, NOT grouped (its effect sits in the middle)
    T = box("T_trigger", 16, 600, 40, text="go")  # click-trigger button, NOT grouped

    q = sl.TimeLine.MainSequence
    e = q.AddEffect(A, FLY, 0, ON_CLICK); e.EffectParameters.Direction = 4; e.Timing.Duration = 0.75
    e = q.AddEffect(B, WIPE, 0, AFTER_PREV); e.EffectParameters.Direction = 11; e.Timing.TriggerDelayTime = 0.5
    q.AddEffect(C, FADE, 2, ON_CLICK)  # by first-level paragraph -> 3 effects
    e = q.AddEffect(N, ZOOM, 0, ON_CLICK)
    e = q.AddEffect(D, SPIN, 0, WITH_PREV); e.EffectParameters.Amount = 180
    e.Timing.RepeatCount = 2; e.Timing.AutoReverse = TRUE
    e = q.AddEffect(E, PATH_DOWN, 0, ON_CLICK)
    e.Behaviors.Item(1).MotionEffect.Path = "M 0 0 L 0.1 0.15 L 0.2 0 E"
    e = q.AddEffect(F, FADE, 0, ON_CLICK); e.Exit = TRUE; e.Timing.Accelerate = 0.3; e.Timing.Decelerate = 0.3
    e = q.AddEffect(G, GROW, 0, AFTER_PREV); e.EffectParameters.Size = 150
    e = q.AddEffect(H, FILL_COLOR, 0, WITH_PREV); e.EffectParameters.Color2.RGB = 0x0000FF
    e = q.AddEffect(J, APPEAR, 0, ON_CLICK)
    e = q.ConvertToAfterEffect(e, 1, 0x808080)  # dim to grey after
    e = q.AddEffect(K, FADE, 0, ON_CLICK)
    e = q.ConvertToTextUnitEffect(e, 1)  # by letter
    iq = sl.TimeLine.InteractiveSequences.Add()
    e = iq.AddEffect(I, ZOOM, 0, ON_SHAPE)
    e.Timing.TriggerShape = T
    pres.SaveAs(path)
    ids = {k: v.Id for k, v in shapes.items()}
    retry(pres.Close)
    return ids


def main(outdir: str):
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    src = str(out / "stress_source.pptx")
    for f in (src, str(out / "stress_source_accessible.pptx")):
        if os.path.exists(f):
            os.remove(f)
    ids = COM.run(build, src)
    print("built", ids)

    def run():
        info = session.open_working_copy(src, fresh=True)
        pres = session.get_pres(info["working_copy"])
        members = [ids[k] for k in ("A_fly", "B_wipe", "C_paragraphs", "D_spin", "E_path", "F_exit", "G_grow",
                                    "H_color", "I_interactive", "J_dim", "K_letters")]
        before = anim.summarize(anim.snapshot(pres.Slides(1)))
        res = ops.group_shapes(pres, 1, members, alt_text="Stress-test diagram.")
        pres = session.get_pres(info["working_copy"])
        after = anim.summarize(anim.snapshot(pres.Slides(1)))
        retry(pres.Close)
        return res, before, after

    res, before, after = COM.run(run)
    print(json.dumps({k: v for k, v in res.items() if k not in ("visual_check",)}, indent=1, default=str)[:6000])
    print("effects before:", len(before), "after:", len(after))
    ok = res.get("ok") and res["xml_check"]["status"] in ("identical", "equivalent")
    print("RESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="ppta11y_stress_")))
