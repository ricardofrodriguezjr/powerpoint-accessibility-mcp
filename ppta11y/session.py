"""Which presentation we are editing, and the safety rules around saving it.

Rules enforced here (not just documented):
  * Every edit happens on a registered *working copy*; the source deck is never saved over.
  * A working copy is saved before any risky operation, so a failed operation can be rolled back by
    reopening the file from disk ("transaction" = last save).
"""
from __future__ import annotations

import json
import os
import shutil
import time
from pathlib import Path

from .com import TRUE, ToolError, app, retry, safe

# Bookkeeping lives next to the deck (<stem>_a11y_checks/working_copy.json), never in AppData:
# PowerPoint runs outside this process's sandbox and cannot see virtualised AppData writes.


def _norm(p: str | os.PathLike) -> str:
    # realpath resolves junctions, symlinks and 8.3 short names for existing files.
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(p))))


def same_file(a: str | os.PathLike, b: str | os.PathLike) -> bool:
    try:
        if os.path.exists(a) and os.path.exists(b):
            return os.path.samefile(a, b)
    except OSError:
        pass
    return _norm(a) == _norm(b)


def checks_dir(deck_path: str | os.PathLike, create: bool = True) -> Path:
    p = Path(deck_path)
    d = p.with_name(p.stem + "_a11y_checks")
    if create:
        d.mkdir(exist_ok=True)
    return d


def _sidecar(deck_path: str) -> Path:
    return checks_dir(deck_path, create=False) / "working_copy.json"


def _read_sidecar(deck_path: str) -> dict | None:
    try:
        return json.loads(_sidecar(deck_path).read_text(encoding="utf8"))
    except Exception:
        return None


def is_working_copy(path: str) -> bool:
    sc = _read_sidecar(path)
    return bool(sc) and same_file(sc.get("working_copy", ""), path)


def source_of(path: str) -> str | None:
    sc = _read_sidecar(path)
    return sc.get("source") if sc and is_working_copy(path) else None


def is_protected(path: str) -> bool:
    """A deck is protected if any working copy in its folder names it as the source."""
    folder = Path(path).parent
    for sc in folder.glob("*_a11y_checks/working_copy.json"):
        try:
            if same_file(json.loads(sc.read_text(encoding="utf8")).get("source", ""), path):
                return True
        except Exception:
            continue
    return False


def find_open(path: str):
    for p in app().Presentations:
        if same_file(p.FullName, path):
            return p
    return None


def open_working_copy(source: str, dest: str | None = None, fresh: bool = False) -> dict:
    """Copy `source` to `dest` (default: <stem>_accessible.pptx next to it), register it, open it."""
    src = Path(source).resolve()
    if not src.exists():
        raise ToolError(f"source deck not found: {src}")
    if src.suffix.lower() not in (".pptx", ".pptm"):
        raise ToolError("only .pptx/.pptm decks are supported")
    out = Path(dest).resolve() if dest else src.with_name(f"{src.stem}_accessible{src.suffix}")
    if same_file(out, src):
        raise ToolError("the working copy must be a different file from the source deck")
    if out.suffix.lower() != src.suffix.lower():
        raise ToolError(f"the working copy must keep the {src.suffix} extension")
    if is_protected(str(out)):
        raise ToolError(f"{out} is registered as a protected original; pick another destination")
    if is_working_copy(str(src)):
        raise ToolError(f"{src} is itself a working copy; open it with open_deck instead")
    if out.exists():
        if not is_working_copy(str(out)):
            raise ToolError(f"{out} already exists and is not a working copy made by this tool. Pick another "
                            "dest (or move that file away yourself); it will not be adopted or overwritten.")
        if not same_file(source_of(str(out)) or "", src):
            raise ToolError(f"{out} is a working copy of a different deck ({source_of(str(out))})")

    other = (_read_sidecar(str(out)) or {}).get("working_copy")
    if other and Path(other).exists() and not same_file(other, out) and same_file(Path(other).parent, out.parent):
        # the checks folder is named by stem, so deck.pptx and deck.pptm in one folder would share it
        raise ToolError(f"{checks_dir(out, create=False)} already belongs to the working copy {other}; pick a "
                        "dest with a different name so the two copies do not share their records")
    note = "resumed existing working copy"
    already = find_open(str(out))
    if fresh or not out.exists():
        if already is not None:
            raise ToolError(f"{out.name} is open in PowerPoint; close it before recreating it")
        backup = checks_dir(out) / f"backup_{time.strftime('%Y%m%d-%H%M%S')}{out.suffix}"
        note_backup = ""
        if out.exists():  # fresh=True on an existing working copy: keep the old one, never overwrite it
            shutil.move(str(out), str(backup))
            note_backup = f"; previous working copy moved to {backup}"
        # '@' id mappings and accepted reading orders belong to the old copy (also when it was deleted)
        for name, suffix in (("id_map.json", ".id_map.json"), ("reading_order_accepted.json",
                                                               ".reading_order_accepted.json")):
            f = checks_dir(out) / name
            if f.exists():
                shutil.move(str(f), str(backup.with_suffix(suffix)))
                note_backup += f"; old {name} moved aside"
        shutil.copy2(src, out)
        note = "created working copy (fresh copy of the source)" + note_backup

    checks_dir(out).joinpath("working_copy.json").write_text(json.dumps(
        {"working_copy": str(out), "source": str(src), "created": time.strftime("%Y-%m-%d %H:%M:%S"),
         "note": "Written by ppta11y. Only the working copy is ever saved; the source is never modified."},
        indent=1), encoding="utf8")

    pres = already or app().Presentations.Open(str(out), 0, 0, TRUE)
    return {"note": note, "working_copy": str(out), "source": str(src), **deck_summary(pres)}


def open_deck(path: str, read_only: bool = False) -> dict:
    """Open (or focus) a deck. Edits are only allowed if it is a registered working copy."""
    p = Path(path).resolve()
    if not p.exists():
        raise ToolError(f"not found: {p}")
    pres = find_open(str(p)) or app().Presentations.Open(str(p), TRUE if read_only else 0, 0, TRUE)
    return {"working_copy": is_working_copy(str(p)), "protected": is_protected(str(p)), **deck_summary(pres)}


def deck_summary(pres) -> dict:
    return {
        "path": pres.FullName,
        "slides": pres.Slides.Count,
        "size_in": [round(pres.PageSetup.SlideWidth / 72, 2), round(pres.PageSetup.SlideHeight / 72, 2)],
        "read_only": bool(pres.ReadOnly),
        "unsaved_changes": not bool(pres.Saved),
    }


def get_pres(deck: str | None = None):
    """Resolve the presentation a tool should act on."""
    a = app()
    if deck:
        p = find_open(deck)
        if p is None:
            raise ToolError(f"{deck} is not open in PowerPoint (use open_working_copy or open_deck)")
        return p
    wc = [p for p in a.Presentations if is_working_copy(p.FullName)]
    if len(wc) == 1:
        return wc[0]
    if a.Presentations.Count == 0:
        raise ToolError("no presentation is open; call open_working_copy first")
    if len(wc) > 1:
        raise ToolError("several working copies are open; pass deck=<full path>: "
                        + "; ".join(p.FullName for p in wc))
    return a.ActivePresentation


def require_editable(pres) -> None:
    path = pres.FullName
    if is_protected(path):
        raise ToolError(f"{path} is a protected original. Edit its working copy instead.")
    if not is_working_copy(path):
        raise ToolError(f"{path} is not a registered working copy. Call open_working_copy on it first "
                        "(this keeps the original untouched).")
    if pres.ReadOnly:
        raise ToolError(f"{path} is open read-only")


def save(pres) -> None:
    require_editable(pres)
    retry(pres.Save)
    for _ in range(60):  # Save() can return before the file is flushed on large decks
        if pres.Saved:
            break
        time.sleep(0.25)


def checkpoint(pres) -> None:
    """Make the file on disk equal the in-memory deck so revert() can restore this exact state."""
    require_editable(pres)
    if not pres.Saved:
        save(pres)


def revert(pres):
    """Discard unsaved changes by closing without saving and reopening from disk. Returns the new object."""
    require_editable(pres)
    path = pres.FullName
    pres.Saved = TRUE
    retry(pres.Close)
    return app().Presentations.Open(path, 0, 0, TRUE)


def list_open() -> list[dict]:
    return [{"path": p.FullName, "working_copy": is_working_copy(p.FullName), "protected": is_protected(p.FullName),
             "unsaved_changes": not bool(p.Saved), "read_only": bool(p.ReadOnly)} for p in app().Presentations]


def close(pres, save_first: bool = False, discard: bool = False) -> None:
    if save_first:
        save(pres)
    elif not pres.Saved and not discard:
        what = "call save_deck first or revert_to_saved" if is_working_copy(pres.FullName) else             "this is not a working copy, so it cannot be saved here; close it in PowerPoint yourself or pass discard=True"
        raise ToolError(f"{pres.Name} has unsaved changes; {what}")
    if discard:
        pres.Saved = TRUE
    retry(pres.Close)


def save_copy(pres, path: str) -> str:
    """Write the in-memory deck to `path` without changing which file the window edits."""
    if same_file(path, pres.FullName) or is_protected(path) or is_working_copy(path):
        raise ToolError("refusing to SaveCopyAs over the working copy or a protected original")
    retry(lambda: pres.SaveCopyAs(path))
    return path


def out_dir(pres) -> Path:
    """Where snapshots and reports for this deck go: <deck folder>/<deck stem>_a11y_checks/."""
    return checks_dir(pres.FullName)


def tmp_dir(pres) -> Path:
    d = out_dir(pres) / "tmp"
    d.mkdir(exist_ok=True)
    return d


def stamp() -> str:
    return time.strftime("%Y%m%d-%H%M%S")


def safe_name(pres) -> str:
    return safe(lambda: pres.Name, "deck")
