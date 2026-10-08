# Changelog

## 2.0.0 (2026-10-08)

First public release as **PowerPoint Accessibility MCP**, forked from
[Ayushmaniar/powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp).

### Added

- Full WCAG 2.1 AA remediation, not only alt text:
  - slide titles (visible heading promoted with an identical look, hidden titles for repeated banners,
    "(n of N)" for continued topics);
  - screen-reader reading order, fixed by restacking objects without any visible change;
  - text contrast measured against the rendered background, in every animation step, fixed with the
    nearest passing colour;
  - tables: diagram tables become identical-looking text boxes, data tables are unmerged and get header
    rows;
  - link text that describes the destination and keeps the link;
  - `wcag_audit`: every rule plus PowerPoint's own Accessibility Checker, with `done: true` when nothing is
    left;
  - a change log with before and after values for every change.
- One-command installer for Claude Code, Codex and OpenCode (`install.ps1`), and `uninstall.ps1`.
- Packaging (`pyproject.toml`): `python -m ppta11y`, `ppt-a11y-mcp` and `ppt-a11y` entry points.

### Changed

- The upstream general-purpose PowerPoint automation server was replaced by the `ppt-a11y` server, which is
  focused on accessibility and on keeping decks visually and functionally identical.

## 1.0.0 (2026-10-02, not published)

- Alt text and "Mark as decorative" for every object, including Slide Master content.
- Grouping multi-piece figures without losing their animations, verified against the saved animation XML
  and by Slide Show playback.
- Applying alt text from a spreadsheet, with results written to a copy.
