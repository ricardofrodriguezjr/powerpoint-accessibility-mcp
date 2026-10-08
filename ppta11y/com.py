"""COM plumbing: one dedicated STA thread for every PowerPoint call, retries, constants."""
from __future__ import annotations

import concurrent.futures
import functools
import threading
import time

import pythoncom
import win32com.client
import win32com.client.dynamic


class ToolError(Exception):
    """An error the caller should see as a plain message (not a traceback)."""


# HRESULTs PowerPoint returns while it is busy (modal dialog, slideshow starting, autosave...).
_BUSY = {-2147418111, -2147417846}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


class _ComThread:
    """All COM objects live on one thread. MCP tool handlers run in a thread pool, so every call is
    marshalled here; PowerPoint COM objects must never cross threads."""

    def __init__(self) -> None:
        self._ex = concurrent.futures.ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="ppt-com", initializer=pythoncom.CoInitialize)

    def run(self, fn, *args, **kwargs):
        # No automatic re-run here: re-running a multi-step edit after a partial failure is unsafe.
        # Individual idempotent calls (Save, Close, SaveCopyAs) use retry() below.
        if threading.current_thread().name.startswith("ppt-com"):
            return fn(*args, **kwargs)  # already on the COM thread (nested call): run inline, no deadlock
        return self._ex.submit(fn, *args, **kwargs).result()


def is_busy(exc: Exception) -> bool:
    return isinstance(exc, pythoncom.com_error) and exc.hresult in _BUSY  # type: ignore[attr-defined]


COM = _ComThread()


def on_com_thread(fn):
    """Decorator: run the function on the COM thread (safe to call from any thread)."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        return COM.run(fn, *args, **kwargs)
    return wrapper


def app():
    # Late binding on purpose: no dependency on the gen_py cache (a common source of breakage).
    a = win32com.client.dynamic.Dispatch("PowerPoint.Application")
    try:
        if not a.Visible:
            a.Visible = -1
    except pythoncom.com_error:  # type: ignore[attr-defined]
        pass  # PowerPoint refuses Visible changes in some states; harmless
    return a


def retry(fn, tries: int = 40, delay: float = 0.5):
    """PowerPoint sometimes rejects Close()/Save() right after a save; retry briefly."""
    last = None
    for _ in range(tries):
        try:
            return fn()
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            last = exc
            time.sleep(delay)
    raise last  # type: ignore[misc]


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def com_err(exc: Exception) -> str:
    if isinstance(exc, pythoncom.com_error):  # type: ignore[attr-defined]
        try:
            return exc.excepinfo[2].strip() or str(exc)
        except Exception:
            return str(exc)
    return str(exc)


TRUE, FALSE = -1, 0

SHAPE_TYPES = {
    -2: "mixed", 1: "autoshape", 2: "callout", 3: "chart", 4: "comment", 5: "freeform", 6: "group",
    7: "embedded_ole", 8: "form_control", 9: "line", 10: "linked_ole", 11: "linked_picture",
    12: "ole_control", 13: "picture", 14: "placeholder", 15: "wordart", 16: "media", 17: "textbox",
    18: "script", 19: "table", 20: "canvas", 21: "diagram", 22: "ink", 23: "ink_comment", 24: "smartart",
    25: "slicer", 26: "web_video", 27: "content_app", 28: "graphic", 29: "linked_graphic",
    30: "3d_model", 31: "linked_3d_model",
}

# Placeholder types (PpPlaceholderType) that hold slide text a screen reader reads anyway.
TEXT_PLACEHOLDERS = {1, 2, 3, 4, 5, 6, 13, 14, 15, 16}  # title, body, ctrTitle, subTitle, vertical*, footer...
TITLE_PLACEHOLDERS = {1, 3, 5}  # ppPlaceholderTitle, CenterTitle, VerticalTitle

# Object kinds PowerPoint's Accessibility Checker flags for missing alt text.
NEEDS_ALT_TYPES = {3, 6, 7, 10, 11, 13, 16, 20, 21, 22, 24, 26, 28, 29, 30, 31}
