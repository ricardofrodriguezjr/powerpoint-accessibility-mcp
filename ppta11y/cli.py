"""Call any ppt-a11y MCP tool from the command line (same code path as the MCP server).

    python cli.py list
    python cli.py inspect_slide "{\\"slide\\": 9}"
    python cli.py group_shapes @args.json        (arguments from a JSON file)
"""
from __future__ import annotations

import inspect
import json
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
from . import server  # noqa: E402

TOOLS = {n: f for n, f in vars(server).items()
         if callable(f) and getattr(f, "__wrapped__", None) is not None and f.__module__ == server.__name__}


def main(argv: list[str]) -> int:
    if len(argv) < 2 or argv[1] in ("-h", "--help", "list"):
        for n, f in sorted(TOOLS.items()):
            print(f"{n}{inspect.signature(f)}\n    {(f.__doc__ or '').strip().splitlines()[0]}")
        return 0
    name = argv[1]
    if name not in TOOLS:
        print(f"unknown tool {name}; try 'list'")
        return 2
    raw = argv[2] if len(argv) > 2 else "{}"
    if raw.startswith("@"):
        raw = Path(raw[1:]).read_text(encoding="utf-8-sig")
    out = TOOLS[name](**json.loads(raw))
    parts = out if isinstance(out, list) else [out]
    for p in parts:
        if isinstance(p, str):
            sys.stdout.buffer.write((p + "\n").encode("utf8"))
        else:
            print(f"[image] {getattr(p, 'path', '')}")
    first = parts[0] if isinstance(parts[0], str) else ""
    return 1 if first.startswith("error:") or '"ok": false' in first else 0


def entry() -> None:
    """Console-script entry point (`ppt-a11y`)."""
    sys.exit(main(sys.argv))
