# Troubleshooting

## Installing

**"Running scripts is disabled on this system"**
The one-line command (`irm ... | iex`) is not affected by this setting. If you downloaded `install.ps1`
and run it as a file, start it like this instead:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

**"Desktop PowerPoint was not found"**
The tool drives the PowerPoint app installed on your computer. Install Microsoft PowerPoint (Microsoft 365
or Office) and run the installer again. PowerPoint on the web and PowerPoint for Mac cannot be automated.

**"could not replace ... Close Claude Code, Codex and OpenCode"**
An open app keeps the server running from the install folder. Quit every app that uses it (including the
desktop apps), then run the installer again.

**"No app was set up"**
The installer looks for the `claude`, `codex` and `opencode` commands. Install the app first, open a new
PowerShell window so it is on the PATH, and run the installer again. You can also name the apps:
`-Apps claude,codex,opencode`. See [Manual setup](manual-setup.md) to configure any other app by hand.

**Behind a university proxy or firewall**
The installer downloads from `github.com`, `astral.sh` (uv), `pypi.org` and `files.pythonhosted.org`, and
uv downloads Python from `github.com`. Ask IT to allow those, or install offline: download the repository
ZIP on another computer, copy it over and run `.\install.ps1 -Source C:\path\to\the.zip`.

## Connecting

**The app does not show the tools**
Restart the app or start a new session after installing. Then check:

| App | Check |
|---|---|
| Claude Code | `claude mcp list`, or type `/mcp` in a session |
| Codex | `codex mcp list`, or type `/mcp` in a session |
| OpenCode | `opencode mcp list` |

If the server shows as failed, run it by hand to see the error:

```powershell
& "$env:USERPROFILE\.ppt-a11y\app\.venv\Scripts\python.exe" "$env:USERPROFILE\.ppt-a11y\app\cli.py" list
```

It should print 46 tools. If it prints an error, run the installer again.

**Codex has no PowerPoint tools in its first answer**
Codex only waits a moment for MCP servers when a session starts. The installer adds
`mcp_optional_startup_grace_ms = 5000` at the top of `%USERPROFILE%\.codex\config.toml` so it waits long
enough; if you set Codex up by hand, add that line yourself. Sending the message again also works.

**"running scripts is disabled" when typing `codex` or `opencode`**
npm installs these apps with a small PowerShell script that Windows blocks by default. Type `codex.cmd` (or
`opencode.cmd`) instead, or run the command in Command Prompt.

**Codex or OpenCode runs inside WSL**
The server must run on Windows itself, because it talks to PowerPoint. Run the app in a normal PowerShell
or Windows Terminal tab, not in WSL.

## While it works

**PowerPoint asks questions or shows a dialog**
Answer or close the dialog; the tool waits for PowerPoint. Leave PowerPoint open and do not work in the
`_accessible` copy while the app is using it.

**"the slide looks different" and the step was undone**
That is the safety net working: every change is checked against a render of the slide and undone if
anything visible changed that should not. The app reports what it could not do invisibly and asks you.

**Animation checks need the screen**
To prove animations still play exactly as before, the app plays slides in Slide Show and captures the
screen. It asks first. Leave the computer alone for those few minutes and keep PowerPoint on the main
monitor.

**PowerPoint's Accessibility Checker still says "Check reading order"**
That item is a prompt to double-check, not an error. It can stay on for a slide whose order is correct
(for example a code slide where note arrows point into the code). The report lists those slides and why
their order is right.

## Removing it

```powershell
irm https://raw.githubusercontent.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/main/uninstall.ps1 | iex
```

This removes the server from every app, the skill copies the installer made, and the install folder. Your
decks and their `_accessible` copies are never touched.
