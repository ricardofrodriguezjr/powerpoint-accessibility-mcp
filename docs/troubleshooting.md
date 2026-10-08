# Troubleshooting

## Installing

**"Running scripts is disabled on this system"**
The one-line command (`irm ... | iex`) is not affected by this setting. If you downloaded `install.ps1`
and run it as a file, start it like this instead:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

**"Desktop PowerPoint was not found"**
The tool drives the PowerPoint app installed on your computer. Install desktop PowerPoint from Microsoft 365
and run the installer again. PowerPoint on the web and PowerPoint for Mac are not supported.

**"the tool is in use" or "could not replace ..."**
An open app keeps the server running from the install folder. Quit Claude Code, Codex and OpenCode
(including their desktop apps and editor extensions), then run the installer again. If an update fails
halfway, the installer puts the previous version back.

**"this window runs as ..., but ... is signed in"**
The PowerShell window was opened with another account (for example an IT administrator account). Open
PowerShell normally, not with "Run as administrator", and run the command again. No administrator rights
are needed.

**"No app was set up"**
The installer finds Claude Code and OpenCode through their commands (`claude`, `opencode`), and Codex through
its command or its settings folder. If you installed an app's command-line version, open a new PowerShell
window and run the installer again. If you use only the Claude desktop app, install the Claude Code command
too (`irm https://claude.ai/install.ps1 | iex`) or follow [Manual setup](manual-setup.md). The server itself
is already installed in `%USERPROFILE%\.ppt-a11y\app`.

**"invalid peer certificate" or "UnknownIssuer"**
Your network inspects secure connections. The installer retries with the Windows certificate store; if it
still fails, ask IT to allow the sites below without inspection.

**Behind a university proxy or firewall**
The installer needs to reach `raw.githubusercontent.com`, `github.com`, `codeload.github.com`, GitHub's
download hosts (`objects.githubusercontent.com`, `release-assets.githubusercontent.com`), `astral.sh` (uv),
`pypi.org` and `files.pythonhosted.org` (uv also downloads Python from GitHub). Ask IT to allow those. If
only the code download is blocked, download the
[ZIP](https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/archive/refs/heads/main.zip) another
way, extract it, and run:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\path\to\powerpoint-accessibility-mcp-main\install.ps1" -Source "C:\path\to\powerpoint-accessibility-mcp-main"
```

The installer still needs the other sites to get uv, Python and the Python packages.

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

**"not a registered working copy" or "is not open in PowerPoint" for a deck in OneDrive or SharePoint**
When Office syncs a file through OneDrive or SharePoint, PowerPoint reports it by its web address, which the
tool cannot match to the file on your disk. Copy the deck to a local folder such as `C:\Slides` and run it
there. Alternatively, in OneDrive settings turn off **Use Office applications to sync Office files that I
open**.

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

**"the Accessibility pane did not open" or the PowerPoint checker was skipped**
The final check reads PowerPoint's Accessibility Checker (the Microsoft 365 Accessibility Assistant) through
its English interface. With PowerPoint in another language or an older Office, the tool's own audit still
runs; at the end, run **Review > Check Accessibility** yourself once.

**PowerPoint's Accessibility Checker still says "Check reading order"**
That item is a prompt to double-check, not an error. It can stay on for a slide whose order is correct
(for example a code slide where note arrows point into the code). The report lists those slides and why
their order is right.

**Claude Code asks before every PowerPoint tool**
Answer **Yes, and don't ask again** the first time each tool is used, or allow them all at once: type
`/permissions` in a session and add an allow rule for `mcp__ppt-a11y`.

## Removing it

```powershell
irm https://raw.githubusercontent.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/main/uninstall.ps1 | iex
```

This removes the server from every app, the skill copies the installer made, and the tool's files in
`%USERPROFILE%\.ppt-a11y`, including the settings backups in `backups`. To keep that folder (and the
backups), download `uninstall.ps1` and run `powershell -ExecutionPolicy Bypass -File .\uninstall.ps1 -KeepFiles`.
uv stays installed; see [uv's uninstall steps](https://docs.astral.sh/uv/getting-started/installation/#uninstallation)
if you no longer need it. Your decks and their `_accessible` copies are never touched.
