"""Play one slide of two decks in a windowed slideshow and grab a frame after each click, so the
animation builds can be compared side by side (proof that grouping kept the animations).

    .venv\\Scripts\\python.exe tests\\playback_compare.py A.pptx B.pptx SLIDE CLICKS OUT.png [WAIT_S]
"""
from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import win32gui
from PIL import Image, ImageChops, ImageDraw, ImageGrab

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ppta11y.com import COM, app, retry  # noqa: E402

ctypes.windll.user32.SetProcessDPIAware()


def frames(path: str, slide: int, clicks: int, wait: float) -> list[Image.Image]:
    pres = app().Presentations.Open(path, -1, 0, -1)
    ss = pres.SlideShowSettings
    ss.StartingSlide, ss.EndingSlide, ss.RangeType, ss.ShowType = slide, slide, 2, 2
    w = ss.Run()
    time.sleep(1.5)
    out = []
    for k in range(clicks + 1):
        if k:
            w.View.Next()
            time.sleep(wait)
        found = []
        win32gui.EnumWindows(lambda h, _: found.append(h) if win32gui.GetWindowText(h).startswith(
            "PowerPoint Slide Show") and win32gui.IsWindowVisible(h) else None, None)
        if not found:
            break
        try:
            win32gui.SetForegroundWindow(found[0])
        except Exception:
            pass
        time.sleep(0.4)
        img = ImageGrab.grab(win32gui.GetWindowRect(found[0]), all_screens=True)
        out.append(img.crop((8, 40, img.width - 8, img.height - 8)))
    w.View.Exit()
    retry(pres.Close)
    return out


def main(a, b, slide, clicks, out, wait=6.0):
    A = COM.run(frames, a, int(slide), int(clicks), float(wait))
    B = COM.run(frames, b, int(slide), int(clicks), float(wait))
    W = 700
    A = [f.resize((W, round(W * f.height / f.width))) for f in A]
    B = [f.resize(A[0].size) for f in B]
    H = A[0].height
    sheet = Image.new("RGB", (W * 2 + 30, (H + 36) * len(A) + 30), "white")
    d = ImageDraw.Draw(sheet)
    d.text((10, 8), f"ORIGINAL  slide {slide}", fill=(0, 0, 0))
    d.text((W + 30, 8), f"GROUPED + ALT TEXT  slide {slide}", fill=(0, 0, 0))
    diffs = []
    for k, (x, y) in enumerate(zip(A, B)):
        top = 30 + k * (H + 36)
        d.text((10, top), f"after click {k}" if k else "slide start", fill=(90, 90, 90))
        sheet.paste(x, (0, top + 16))
        sheet.paste(y, (W + 30, top + 16))
        bb = ImageChops.difference(x.convert("L"), y.convert("L")).point(lambda v: 255 if v > 40 else 0).getbbox()
        diffs.append(bb)
    sheet.save(out)
    print("frames:", len(A), "differences per frame:", diffs)


if __name__ == "__main__":
    main(*sys.argv[1:])
