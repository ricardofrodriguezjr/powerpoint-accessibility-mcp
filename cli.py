#!/usr/bin/env python3
r"""Call any ppt-a11y tool from a terminal, without an AI app (same code path as the MCP server).

    .venv\Scripts\python.exe cli.py list
    .venv\Scripts\python.exe cli.py wcag_audit "@args.json"      (arguments from a JSON file)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ppta11y.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main(sys.argv))
