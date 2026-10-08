# Tests

These scripts drive the real PowerPoint on your computer, so they need Windows and desktop PowerPoint.
They only ever work on decks they create, or on copies in the scratch folder you give them. Close other
PowerPoint windows first, and do not use the computer while the playback tests run.

Run them from the repository folder after setting up the environment
([Manual setup](../docs/manual-setup.md)):

| Command | What it proves |
|---|---|
| `.venv\Scripts\python.exe tests\smoke_mcp.py` | The server starts over stdio, answers the MCP handshake and lists all 46 tools. |
| `.venv\Scripts\python.exe tests\wcag_v2_tests.py C:\Temp\ppt-test` | 27 checks (4 on their own, 23 end to end on a synthetic deck): titles, unique titles, tables, links, contrast (including rotated text and an animated cover), reading order, and that nothing visible changed except what WCAG requires. Prints `RESULT: PASS`. |
| `.venv\Scripts\python.exe tests\stress_animations.py C:\Temp\ppt-stress` | Groups 11 shapes carrying 13 kinds of animation and requires PowerPoint's saved animation XML to be identical afterwards (or equivalent: the same animation with only internal ids or order changed). Prints `RESULT: PASS`. |
| `.venv\Scripts\python.exe tests\deck_diff.py ORIGINAL.pptx COPY.pptx C:\Temp\diff` | Exports every slide of two decks and lists the slides that differ (beyond antialiasing). Useful after a real remediation. |
| `.venv\Scripts\python.exe tests\playback_all.py ORIGINAL.pptx COPY.pptx C:\Temp\playback.json` | Plays every animated slide of both decks in Slide Show and compares a frame after each click. |
| `.venv\Scripts\python.exe tests\playback_compare.py A.pptx B.pptx SLIDE CLICKS out.png` | Side-by-side frames of one slide, click by click. |
