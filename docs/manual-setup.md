# Manual setup

The [one-line installer](../README.md#quick-start) does all of this for you. Use this page if you prefer to
see every step, want to install somewhere else, or use an MCP app other than Claude Code, Codex and
OpenCode.

Run every command in **Windows PowerShell** (Start menu > type `PowerShell`). Do not use WSL: the server
must run on Windows itself to reach PowerPoint.

## 1. Get the code and a Python environment

Install [uv](https://docs.astral.sh/uv/) once. It also downloads a private Python for this project, so you
do not need to install Python yourself:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Open a new PowerShell window, then download the code and create the environment:

```powershell
git clone https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp "$env:USERPROFILE\.ppt-a11y\app"
cd "$env:USERPROFILE\.ppt-a11y\app"
uv venv .venv --python 3.12
uv pip install --python .venv\Scripts\python.exe -r requirements.txt
.venv\Scripts\python.exe cli.py list
```

No git? Use these lines instead of the `git clone` line, then continue with the `cd` line:

```powershell
Invoke-WebRequest https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/archive/refs/heads/main.zip -OutFile "$env:TEMP\ppt-a11y.zip"
Expand-Archive "$env:TEMP\ppt-a11y.zip" "$env:USERPROFILE\.ppt-a11y" -Force
Rename-Item "$env:USERPROFILE\.ppt-a11y\powerpoint-accessibility-mcp-main" app
```

The last command of the first block should print 46 tools.

The rest of this page uses these two paths. Adjust them if you put the code somewhere else:

| What | Path |
|---|---|
| Python | `%USERPROFILE%\.ppt-a11y\app\.venv\Scripts\python.exe` |
| Server | `%USERPROFILE%\.ppt-a11y\app\server.py` |

## 2. Connect your app

<details open>
<summary><b>Claude Code</b> (CLI, desktop app, IDE extensions)</summary>

Register the server for your user account:

```powershell
claude mcp add --scope user ppt-a11y -- "$env:USERPROFILE\.ppt-a11y\app\.venv\Scripts\python.exe" "$env:USERPROFILE\.ppt-a11y\app\server.py"
```

Install the skill:

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null
Copy-Item -Recurse -Force "$env:USERPROFILE\.ppt-a11y\app\skills\pptx-wcag-alt-text" "$env:USERPROFILE\.claude\skills\"
```

Start a new session and type `/mcp`: **ppt-a11y** should be connected. Type `/pptx-wcag-alt-text` to use the
skill by name, or just ask. A deck needs many tool calls: answer **Yes, and don't ask again** the first time
Claude Code asks to use a PowerPoint tool, or add an allow rule for `mcp__ppt-a11y` with `/permissions`.

</details>

<details>
<summary><b>Codex</b> (CLI, desktop app, IDE extension)</summary>

Register the server:

```powershell
codex mcp add ppt-a11y '--' "$env:USERPROFILE\.ppt-a11y\app\.venv\Scripts\python.exe" "$env:USERPROFILE\.ppt-a11y\app\server.py"
```

Then open `%USERPROFILE%\.codex\config.toml`. The command created a `[mcp_servers.ppt-a11y]` section; add
the last three lines below to it, so that long checks are not cut off and Codex does not ask for approval
before each of the hundreds of tool calls a deck needs (the editing tools only touch the `_accessible`
copy; the `evaluate` tool, which can run code, is off unless you enable it).
Also add `mcp_optional_startup_grace_ms = 5000` as the very first line of the file, so Codex waits for the
server (about 2 seconds to start) before your first message. The finished file looks like this; single
quotes keep Windows backslashes as they are. Do not add a second `[mcp_servers.ppt-a11y]` line: TOML does
not allow the same section twice.

```toml
mcp_optional_startup_grace_ms = 5000

[mcp_servers.ppt-a11y]
command = 'C:\Users\YOUR-NAME\.ppt-a11y\app\.venv\Scripts\python.exe'
args = ['C:\Users\YOUR-NAME\.ppt-a11y\app\server.py']
startup_timeout_sec = 60
tool_timeout_sec = 1800
default_tools_approval_mode = "approve"
```

In the Codex desktop app you can also add the server under **Settings > MCP servers > Add server** (type
STDIO, the Python path as the command, the server path as the argument), then add the extra lines above to
`config.toml`.

Install the skill (Codex reads user skills from `%USERPROFILE%\.agents\skills`):

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.agents\skills" | Out-Null
Copy-Item -Recurse -Force "$env:USERPROFILE\.ppt-a11y\app\skills\pptx-wcag-alt-text" "$env:USERPROFILE\.agents\skills\"
```

Restart Codex and type `/mcp`: **ppt-a11y** should be listed. Type `$pptx-wcag-alt-text` to use the skill by
name, or just ask. In the Codex desktop app and the VS Code extension, keep the agent on **Windows native**
(not WSL).

</details>

<details>
<summary><b>OpenCode</b></summary>

Add the server to `%USERPROFILE%\.config\opencode\opencode.json`. OpenCode creates `opencode.jsonc` in that
folder the first time it runs; if that is the file you have, add the entry there instead (OpenCode reads
both). If the file already has content, keep a single outer `{ }` and add only the `"mcp"` part (or, if it
already has an `"mcp"` section, only the `"ppt-a11y"` entry), with a comma after the entry before it. Use
double backslashes in JSON:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "ppt-a11y": {
      "type": "local",
      "command": [
        "C:\\Users\\YOUR-NAME\\.ppt-a11y\\app\\.venv\\Scripts\\python.exe",
        "C:\\Users\\YOUR-NAME\\.ppt-a11y\\app\\server.py"
      ],
      "enabled": true,
      "timeout": 1800000
    }
  }
}
```

The skill: OpenCode also reads `%USERPROFILE%\.claude\skills` and `%USERPROFILE%\.agents\skills`. If you set
up Claude Code or Codex above, there is nothing else to do. Otherwise:

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.agents\skills" | Out-Null
Copy-Item -Recurse -Force "$env:USERPROFILE\.ppt-a11y\app\skills\pptx-wcag-alt-text" "$env:USERPROFILE\.agents\skills\"
```

Restart OpenCode and run `opencode mcp list`. Run OpenCode in a normal PowerShell or Windows Terminal tab,
not in WSL.

</details>

<details>
<summary><b>Other MCP apps</b></summary>

Any app that supports local (stdio) MCP servers works. Use this command, with no arguments other than the
script path:

```text
C:\Users\YOUR-NAME\.ppt-a11y\app\.venv\Scripts\python.exe C:\Users\YOUR-NAME\.ppt-a11y\app\server.py
```

If the app supports Agent Skills, copy `skills\pptx-wcag-alt-text` into its skills folder. If it does not,
paste the contents of `SKILL.md` into the app's custom instructions.

</details>

## 3. Try it

Open the app and paste a prompt from [Example prompts](example-prompts.md), for example:

```text
Make "C:\Users\me\Documents\STAT 101\Lecture 05.pptx" WCAG 2.1 AA compliant and save the
_accessible copy in the same folder.
```

## Updating

Close your apps first. If you used the installer or the ZIP, run the one-line installer again. If you used
`git clone`:

```powershell
cd "$env:USERPROFILE\.ppt-a11y\app"
git pull
uv pip install --python .venv\Scripts\python.exe -r requirements.txt
```

Then copy the skill folder again (step 2) and restart your app.
