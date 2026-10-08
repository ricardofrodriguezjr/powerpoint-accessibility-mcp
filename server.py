#!/usr/bin/env python3
r"""Start the ppt-a11y MCP server over stdio (this is what your AI app runs).

    .venv\Scripts\python.exe server.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ppta11y.server import main  # noqa: E402

if __name__ == "__main__":
    main()
