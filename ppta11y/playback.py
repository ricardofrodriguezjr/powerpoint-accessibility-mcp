"""End-to-end animation check: play one slide of two decks in a windowed Slide Show, grab a frame after
every click, and compare. This is the only check that looks at what an audience actually sees (it
caught a flattened text box whose fill appeared before its click, which XML/timing checks missed).

Safeguards against a silent pass: the capture must come from the Slide Show window itself (it must be
in front and cover the sampled points), the show must be running on the right slide at every frame,
advance is forced to manual, both decks must need the same number of clicks, frames are captured only
once they stop changing, and the comparison is done at full resolution with a sensitive threshold.

Needs the PowerPoint window to be on screen (it captures the screen). Takes ~1-3 s per click.
"""
from __future__ import annotations

import ctypes
import time
from pathlib import Path

import pythoncom
from PIL import Image, ImageChops, ImageDraw, ImageStat

from . import animations as anim
from .com import TRUE, FALSE, ToolError, app, retry, safe
from .session import find_open, out_dir, stamp

try:  # one DPI context for window coordinates and the screen grab (per-monitor v2)
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

CHANNEL_THRESHOLD = 12  # per-channel difference that counts as "changed" (pale fills differ by ~13-40)


def _ppt_windows() -> set[int]:
    """Visible top-level windows that belong to the PowerPoint process (language independent)."""
    import win32gui
    import win32process
    pids = set()
    out = []
    win32gui.EnumWindows(lambda h, _: out.append(h) if win32gui.IsWindowVisible(h) else None, None)
    for h in out:
        try:
            _, pid = win32process.GetWindowThreadProcessId(h)
        except Exception:
            continue
        if win32gui.GetClassName(h) in ("PPTFrameClass", "screenClass"):
            pids.add(pid)
    return {h for h in out if win32process.GetWindowThreadProcessId(h)[1] in pids}


def _hwnd_of(w, before: set[int]) -> int | None:
    """The Slide Show window: SlideShowWindow.HWND if COM exposes it, else the PowerPoint window that
    appeared when the show started."""
    for flags in (pythoncom.DISPATCH_PROPERTYGET, pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD):
        try:
            o = w._oleobj_
            return int(o.Invoke(o.GetIDsOfNames("HWND"), 0, flags, True)) & 0xFFFFFFFF
        except Exception:
            pass
    for _ in range(20):
        new = _ppt_windows() - before
        if len(new) == 1:
            return new.pop()
        time.sleep(0.25)
    return None


def _slide_box(hwnd, aspect: float) -> tuple[int, int, int, int]:
    """Screen rectangle of the slide inside the show window (client area minus letterbox bars)."""
    import win32gui
    l, t, r, b = win32gui.GetClientRect(hwnd)
    x0, y0 = win32gui.ClientToScreen(hwnd, (l, t))
    cw, ch = r - l, b - t
    w, h = (round(ch * aspect), ch) if cw / ch > aspect else (cw, round(cw / aspect))
    sx, sy = x0 + (cw - w) // 2, y0 + (ch - h) // 2
    return sx + 2, sy + 2, sx + w - 2, sy + h - 2


def _front_and_visible(hwnd, box) -> bool:
    """True if the show window really owns the pixels we are about to capture (sampled inside the slide)."""
    import win32gui
    try:
        win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.25)
    x0, y0, x1, y1 = box
    dx, dy = (x1 - x0) // 10, (y1 - y0) // 10
    pts = [((x0 + x1) // 2, (y0 + y1) // 2), (x0 + dx, y0 + dy), (x1 - dx, y0 + dy), (x0 + dx, y1 - dy),
           (x1 - dx, y1 - dy)]
    return all(win32gui.GetAncestor(win32gui.WindowFromPoint(p), 2) == hwnd for p in pts)  # GA_ROOT


def _grab(hwnd, aspect: float) -> Image.Image:
    """Capture only the slide, once the picture has stopped changing."""
    from PIL import ImageGrab
    box = _slide_box(hwnd, aspect)
    if not _front_and_visible(hwnd, box):
        raise ToolError("the Slide Show window is covered or not in front; keep PowerPoint visible on screen "
                        "(no other window over the slide, screen unlocked) and retry")
    prev = ImageGrab.grab(box, all_screens=True)
    for _ in range(20):  # wait until two consecutive grabs agree (animation finished)
        time.sleep(0.3)
        cur = ImageGrab.grab(box, all_screens=True)
        if ImageChops.difference(prev.convert("RGB"), cur.convert("RGB")).getbbox() is None:
            return cur
        prev = cur
    return prev


def _steps(slide_obj) -> tuple[list[float], float]:
    """Seconds to wait after each click (longest effect end in that click step, plus a margin)."""
    snap = anim.snapshot(slide_obj)
    for q in snap:  # auto-reverse plays the effect twice
        for e in q["effects"]:
            if e["Timing"].get("AutoReverse"):
                e["Timing"]["Duration"] = (e["Timing"].get("Duration") or 0) * 2
    model = anim.timeline_model(snap)
    ends: dict[int, float] = {}
    for entries in model.values():
        for seq, step, start, end, *_ in entries:
            if seq == "main":
                ends[step] = max(ends.get(step, 0.0), end)
    n = max(ends) if ends else 0
    trans = safe(lambda: slide_obj.SlideShowTransition.Duration, 0.0) or 0.0
    return [ends.get(k, 0.0) + 0.6 for k in range(1, n + 1)], ends.get(0, 0.0) + trans + 0.6


def _frames(pres, slide: int, waits: list[float], first_wait: float) -> list[Image.Image]:
    aspect = pres.PageSetup.SlideWidth / pres.PageSetup.SlideHeight
    ss = pres.SlideShowSettings
    sl = pres.Slides(slide)
    keys = ("RangeType", "StartingSlide", "EndingSlide", "ShowType", "AdvanceMode")
    saved = {k: retry(lambda k=k: getattr(ss, k)) for k in keys}
    hidden = sl.SlideShowTransition.Hidden
    was_saved = bool(pres.Saved)
    frames, w = [], None
    try:
        if hidden:
            sl.SlideShowTransition.Hidden = FALSE
        ss.RangeType, ss.StartingSlide, ss.EndingSlide, ss.ShowType, ss.AdvanceMode = 2, slide, slide, 2, 1
        before_windows = _ppt_windows()
        w = ss.Run()
        time.sleep(1.0 + first_wait)
        hwnd = _hwnd_of(w, before_windows)
        if hwnd is None:
            raise ToolError("could not get the Slide Show window handle")
        target = sl.SlideID
        for k in range(len(waits) + 1):
            if k:
                w.View.Next()
                time.sleep(waits[k - 1])
            if safe(lambda: w.View.State) != 1 or safe(lambda: w.View.Slide.SlideID) != target:
                raise ToolError(f"the Slide Show left slide {slide} after {k} click(s) (state "
                                f"{safe(lambda: w.View.State)}); it needs fewer clicks than its animations suggest")
            frames.append(_grab(hwnd, aspect))
    finally:
        if w is not None:
            safe(lambda: w.View.Exit())
            time.sleep(0.3)
        retry(lambda: setattr(ss, "RangeType", saved["RangeType"]))
        for k in ("EndingSlide", "StartingSlide") if (saved["StartingSlide"] or 1) > slide else ("StartingSlide", "EndingSlide"):
            retry(lambda k=k: setattr(ss, k, saved[k]))
        retry(lambda: setattr(ss, "ShowType", saved["ShowType"]))
        retry(lambda: setattr(ss, "AdvanceMode", saved["AdvanceMode"]))
        if hidden:
            sl.SlideShowTransition.Hidden = TRUE
        bad = {k: (getattr(ss, k), v) for k, v in saved.items() if getattr(ss, k) != v}
        if bad:
            raise ToolError(f"could not restore Slide Show settings {bad}; revert the deck before saving")
        if was_saved:
            pres.Saved = TRUE
    return frames


def _diff(a: Image.Image, b: Image.Image) -> tuple[int, tuple | None]:
    if a.size != b.size:
        b = b.resize(a.size, Image.BOX)
    r, g, bb = ImageChops.difference(a.convert("RGB"), b.convert("RGB")).split()
    d = ImageChops.lighter(ImageChops.lighter(r, g), bb).point(lambda v: 255 if v > CHANNEL_THRESHOLD else 0)
    return d.histogram()[255], d.getbbox()


def compare_playback(pres, slide: int, source_deck: str, source_slide: int | None = None,
                     tolerance_px: int = 60) -> dict:
    source_slide = source_slide or slide
    src = find_open(source_deck)
    opened = src is None
    if opened:
        src = app().Presentations.Open(str(Path(source_deck).resolve()), TRUE, FALSE, TRUE)
    try:
        waits_a, first_a = _steps(src.Slides(source_slide))
        waits_b, first_b = _steps(pres.Slides(slide))
        if len(waits_a) != len(waits_b):
            return {"slide": slide, "ok": False, "clicks": [len(waits_a), len(waits_b)], "frames": [],
                    "image": None, "failure": f"click count differs: original {len(waits_a)}, working copy {len(waits_b)}"}
        waits = [max(a, b) for a, b in zip(waits_a, waits_b)]
        first = max(first_a, first_b)
        A = _frames(src, source_slide, waits, first)
    finally:
        if opened:
            retry(src.Close)
    B = _frames(pres, slide, waits, first)
    rows = []
    for k, (a, b) in enumerate(zip(A, B)):
        n, box = _diff(a, b)
        rows.append({"frame": k, "after_click": k, "changed_px": n, "box": box})
    blank = [k for k, a in enumerate(A) if max(ImageStat.Stat(a.convert("L")).stddev) < 1.0]
    width = 640
    As = [f.resize((width, round(width * f.height / f.width)), Image.BOX) for f in A]
    Bs = [f.resize(As[0].size, Image.BOX) for f in B]
    h = As[0].height
    sheet = Image.new("RGB", (width * 2 + 30, (h + 30) * len(rows) + 30), "white")
    dr = ImageDraw.Draw(sheet)
    dr.text((10, 8), f"ORIGINAL slide {source_slide}", fill=(0, 0, 0))
    dr.text((width + 30, 8), f"WORKING COPY slide {slide}", fill=(0, 0, 0))
    sx = width / A[0].width
    for k, (a, b) in enumerate(zip(As, Bs)):
        top = 30 + k * (h + 30)
        bad = rows[k]["changed_px"] > tolerance_px
        dr.text((10, top), ("slide start" if k == 0 else f"after click {k}") + ("   <-- DIFFERENT" if bad else ""),
                fill=(200, 0, 0) if bad else (90, 90, 90))
        sheet.paste(a, (0, top + 14))
        sheet.paste(b, (width + 30, top + 14))
        if bad and rows[k]["box"]:
            x0, y0, x1, y1 = (round(v * sx) for v in rows[k]["box"])
            dr.rectangle([width + 30 + x0, top + 14 + y0, width + 30 + x1, top + 14 + y1], outline=(255, 0, 0), width=3)
    out = out_dir(pres) / f"s{slide}_playback_{stamp()}.png"
    if sheet.height > 4000:
        sheet = sheet.resize((round(sheet.width * 4000 / sheet.height), 4000))
    sheet.save(out)
    ok = (len(A) == len(B) == len(waits) + 1 and not blank and all(r["changed_px"] <= tolerance_px for r in rows))
    res = {"slide": slide, "source": f"{source_deck} slide {source_slide}", "clicks": len(waits),
           "frames": rows, "ok": ok, "image": str(out)}
    if blank:
        res["failure"] = f"frames {blank} of the original are a single colour (nothing was captured)"
    return res
