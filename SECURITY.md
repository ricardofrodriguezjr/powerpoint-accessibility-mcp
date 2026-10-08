# Security and privacy

## What runs where

- Everything runs on your own computer. The MCP server talks only to the PowerPoint installed on your
  computer and to the AI app that started it. It makes no network connections of its own.
- Your AI app (Claude Code, Codex or OpenCode) sends what the tools return (slide text, alt text, and slide
  images when a snapshot is requested) to its AI provider, as with any other file you work on in that app.
  Use the app and provider your institution approves for your course materials.
- The installer downloads this repository from GitHub, a private Python and the Python packages from PyPI
  using [uv](https://docs.astral.sh/uv/) (and installs uv itself to `%USERPROFILE%\.local\bin` if it is
  missing). It writes only to `%USERPROFILE%\.ppt-a11y`, your apps' skills folders and your apps' MCP
  settings, and backs up every settings file before changing it.
- For Codex, the installer turns off the approval prompt for this server's tools only
  (`default_tools_approval_mode = "approve"`), because a deck needs hundreds of tool calls and the editing
  tools only touch the `_accessible` copy. Remove that line
  from `%USERPROFILE%\.codex\config.toml` if you prefer to approve each call. It also adds
  `mcp_optional_startup_grace_ms = 5000` at the top of that file (Codex then waits up to 5 seconds for MCP
  servers when a session starts); the uninstaller removes it again.

## Your files

- The original deck is never saved or modified. All edits go to `<name>_accessible.pptx`.
- Spreadsheets are never edited in place; results go to a copy.
- The `evaluate` tool runs Python code against an open working copy for inspection. It is not a sandbox,
  so it is off unless the server is started with the environment variable `PPT_A11Y_ENABLE_EVALUATE=1`.
  The skill never needs it.
- In Claude Code, an allow rule for `mcp__ppt-a11y` approves every tool of this server at once. With
  `evaluate` off, that covers only tools that read the deck or edit the `_accessible` copy.

## Reporting a vulnerability

Please report security problems privately through
[GitHub security advisories](https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/security/advisories/new)
rather than in a public issue.
