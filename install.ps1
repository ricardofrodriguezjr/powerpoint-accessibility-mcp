<#
.SYNOPSIS
    Installs PowerPoint Accessibility MCP (server "ppt-a11y") and its agent skill, and sets them up for
    Claude Code, Codex and OpenCode.

.DESCRIPTION
    1. Checks for Windows and desktop PowerPoint.
    2. Installs uv (https://docs.astral.sh/uv/) if it is missing. uv provides a private Python.
    3. Downloads this repository to %USERPROFILE%\.ppt-a11y\app and creates a private Python environment.
    4. For every app found (or the ones named with -Apps): registers the MCP server and copies the skill.

    Run it again at any time to update. Nothing outside %USERPROFILE%\.ppt-a11y, the apps' skill folders
    and the apps' MCP settings is changed (plus uv itself, if it had to be installed). Existing settings
    files and skill folders are backed up to %USERPROFILE%\.ppt-a11y\backups before they are changed.

.EXAMPLE
    irm https://raw.githubusercontent.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/main/install.ps1 | iex

.EXAMPLE
    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/main/install.ps1))) -Apps codex

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install.ps1 -Apps claude,opencode
#>
param(
    # Which apps to set up: auto (every app found), claude, codex, opencode, none.
    [string[]]$Apps = @('auto'),
    # Where everything is installed: code, Python environment, private Python, backups. Kept outside
    # AppData on purpose: packaged (MSIX) apps see a private copy of AppData, so files written there from
    # one app can be invisible to another.
    [string]$InstallDir = (Join-Path $env:USERPROFILE '.ppt-a11y'),
    # Branch or tag to download.
    [string]$Ref = 'main',
    # Use a local folder or .zip instead of downloading (offline installs and testing).
    [string]$Source = '',
    # Name the server is registered under in each app.
    [string]$ServerName = 'ppt-a11y',
    # Do not copy the skill.
    [switch]$NoSkill,
    # Home folder that holds the apps' settings (testing only).
    [string]$HomeDir = $env:USERPROFILE
)

$Repo = 'ricardofrodriguezjr/powerpoint-accessibility-mcp'
$MinTools = 40

function Write-Step([string]$Text) { Write-Host ''; Write-Host "==> $Text" -ForegroundColor Cyan }
function Write-Ok([string]$Text) { Write-Host "    [ok] $Text" -ForegroundColor Green }
function Write-Note([string]$Text) { Write-Host "    $Text" }
function Write-Warn2([string]$Text) { Write-Host "    [!] $Text" -ForegroundColor Yellow }

function Invoke-Native {
    # Runs a native program without letting its stderr turn into a PowerShell error.
    param([string]$Exe, [string[]]$Arguments)
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $Exe @Arguments 2>&1 | ForEach-Object { "$_" }
        return [pscustomobject]@{ Code = $LASTEXITCODE; Output = ($out -join "`n") }
    } finally {
        $ErrorActionPreference = $old
    }
}

function Find-Exe([string]$Name) {
    # Programs only (.exe, then .cmd/.bat): npm also installs a .ps1 shim, which a default execution policy
    # blocks, and an extension-less shell script that Windows cannot run.
    $c = Get-Command $Name -CommandType Application -ErrorAction SilentlyContinue |
        Where-Object { @('.exe', '.cmd', '.bat') -contains [IO.Path]::GetExtension($_.Source).ToLower() } |
        Sort-Object { if ($_.Source -like '*.exe') { 0 } else { 1 } } | Select-Object -First 1
    if ($c) { return $c.Source }
    return $null
}

function Backup-Item([string]$Path, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $dir = Join-Path $InstallDir 'backups'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $dest = Join-Path $dir "$Label-$stamp"
    Copy-Item -LiteralPath $Path -Destination $dest -Recurse -Force
    return $dest
}

function Install-Skills([string]$SkillsRoot, [string]$Label) {
    # Copies every skill folder of the repository into $SkillsRoot, backing up a different existing copy.
    $src = Join-Path $script:App 'skills'
    New-Item -ItemType Directory -Force -Path $SkillsRoot | Out-Null
    foreach ($s in Get-ChildItem -LiteralPath $src -Directory) {
        if (-not (Test-Path -LiteralPath (Join-Path $s.FullName 'SKILL.md'))) { continue }
        $dest = Join-Path $SkillsRoot $s.Name
        if (Test-Path -LiteralPath $dest) {
            $b = Backup-Item $dest "$Label-skill-$($s.Name)"
            Remove-Item -LiteralPath $dest -Recurse -Force
            Write-Note "previous copy of the skill saved to $b"
        }
        Copy-Item -LiteralPath $s.FullName -Destination $dest -Recurse -Force
        Set-Content -LiteralPath (Join-Path $dest '.installed-by-ppt-a11y') -Value $script:App -Encoding ASCII
        Write-Ok "skill '$($s.Name)' -> $dest"
    }
}

function Get-SkillNames {
    $src = Join-Path $script:App 'skills'
    return @(Get-ChildItem -LiteralPath $src -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'SKILL.md') } | ForEach-Object { $_.Name })
}

function Get-Code {
    # Puts the repository in $InstallDir\app (replacing an older copy).
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("ppt-a11y-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
        if ($Source -and (Test-Path -LiteralPath $Source -PathType Container)) {
            $root = Join-Path $tmp 'src'
            New-Item -ItemType Directory -Force -Path $root | Out-Null
            Get-ChildItem -LiteralPath $Source -Force | Where-Object { $_.Name -notin @('.git', '.venv', '__pycache__') } |
                ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $root -Recurse -Force }
            Write-Ok "using local folder $Source"
        } else {
            $zip = Join-Path $tmp 'repo.zip'
            if ($Source) {
                Copy-Item -LiteralPath $Source -Destination $zip
                Write-Ok "using local archive $Source"
            } else {
                $kind = 'heads'
                if ($Ref -match '^v\d') { $kind = 'tags' }
                $url = "https://github.com/$Repo/archive/refs/$kind/$Ref.zip"
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
                $pp = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try { Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing } finally { $ProgressPreference = $pp }
                Write-Ok "downloaded $url"
            }
            Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $tmp 'x') -Force
            $root = (Get-ChildItem -LiteralPath (Join-Path $tmp 'x') -Directory | Select-Object -First 1).FullName
        }
        if (-not (Test-Path -LiteralPath (Join-Path $root 'server.py'))) { throw "the download does not contain server.py ($root)" }
        if (Test-Path -LiteralPath $script:App) {
            try {
                Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction Stop
            } catch {
                throw "could not replace $($script:App). Close Claude Code, Codex and OpenCode (they keep the server running), then run the installer again."
            }
        }
        New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
        Move-Item -LiteralPath $root -Destination $script:App
    } finally {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-Uv {
    $uv = Find-Exe 'uv'
    if ($uv) { return $uv }
    foreach ($p in @((Join-Path $HomeDir '.local\bin\uv.exe'), (Join-Path $HomeDir '.cargo\bin\uv.exe'))) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    Write-Note 'uv is not installed; installing it from https://astral.sh/uv (about 20 MB)'
    $r = Invoke-Native 'powershell' @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', 'irm https://astral.sh/uv/install.ps1 | iex')
    if ($r.Code -ne 0) { throw "installing uv failed:`n$($r.Output)" }
    foreach ($p in @((Join-Path $HomeDir '.local\bin\uv.exe'), (Join-Path $HomeDir '.cargo\bin\uv.exe'))) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    $uv = Find-Exe 'uv'
    if ($uv) { return $uv }
    throw 'uv was installed but cannot be found. Open a new PowerShell window and run the installer again.'
}

function Set-Claude {
    $claude = Find-Exe 'claude'
    if (-not $claude) { Write-Warn2 'Claude Code (claude) not found on PATH; skipped'; return $false }
    Invoke-Native $claude @('mcp', 'remove', $ServerName, '--scope', 'user') | Out-Null
    $r = Invoke-Native $claude @('mcp', 'add', '--scope', 'user', $ServerName, '--', $script:Py, $script:Server)
    if ($r.Code -ne 0) { Write-Warn2 "claude mcp add failed:`n$($r.Output)"; return $false }
    Write-Ok "MCP server '$ServerName' registered (user scope)"
    if (-not $NoSkill) { Install-Skills (Join-Path $HomeDir '.claude\skills') 'claude' }
    return $true
}

function Set-Codex {
    $codex = Find-Exe 'codex'
    if (-not $codex) { Write-Warn2 'Codex (codex) not found on PATH; skipped'; return $false }
    Invoke-Native $codex @('mcp', 'remove', $ServerName) | Out-Null
    $r = Invoke-Native $codex @('mcp', 'add', $ServerName, '--', $script:Py, $script:Server)
    if ($r.Code -ne 0) { Write-Warn2 "codex mcp add failed:`n$($r.Output)"; return $false }
    Write-Ok "MCP server '$ServerName' registered"
    Set-CodexOptions
    # Codex reads user skills from ~/.agents/skills (~/.codex/skills is deprecated).
    if (-not $NoSkill) { Install-Skills (Join-Path $HomeDir '.agents\skills') 'codex' }
    return $true
}

function Set-CodexOptions {
    # `codex mcp add` cannot set these, so they are added to the server's table in config.toml:
    # - a longer start-up wait and a 30-minute per-call limit (a full audit or a playback check takes minutes);
    # - no approval prompt for each call of this server's tools (a deck needs hundreds of calls; the tools
    #   only ever edit the _accessible copy).
    $codexHome = Join-Path $HomeDir '.codex'
    if ($env:CODEX_HOME) { $codexHome = $env:CODEX_HOME }
    $cfg = Join-Path $codexHome 'config.toml'
    if (-not (Test-Path -LiteralPath $cfg)) { Write-Warn2 "$cfg not found; options left at their defaults"; return }
    $text = [IO.File]::ReadAllText($cfg)
    $header = "[mcp_servers.$ServerName]"
    if ($text.IndexOf($header) -lt 0) { $header = "[mcp_servers.`"$ServerName`"]" }
    $i = $text.IndexOf($header)
    if ($i -lt 0) { Write-Warn2 "could not find $header in $cfg; options left at their defaults"; return }
    $next = $text.IndexOf("`n[", $i + $header.Length)
    if ($next -lt 0) { $next = $text.Length }
    $block = $text.Substring($i, $next - $i)
    $add = ''
    if ($block -notmatch '(?m)^\s*startup_timeout_sec\s*=') { $add += "startup_timeout_sec = 60`n" }
    if ($block -notmatch '(?m)^\s*tool_timeout_sec\s*=') { $add += "tool_timeout_sec = 1800`n" }
    if ($block -notmatch '(?m)^\s*default_tools_approval_mode\s*=') { $add += "default_tools_approval_mode = `"approve`"`n" }
    # Codex waits only 1 s for MCP servers when it builds the first turn's tool list; this server needs
    # about 2 s to start. A top-level key (it must come before the first [table]).
    $grace = $text -notmatch '(?m)^\s*mcp_optional_startup_grace_ms\s*='
    if (-not $add -and -not $grace) { return }
    Backup-Item $cfg 'codex-config.toml' | Out-Null
    if ($add) {
        $lineEnd = $text.IndexOf("`n", $i)
        if ($lineEnd -lt 0) { $text = $text + "`n" + $add } else { $text = $text.Insert($lineEnd + 1, $add) }
    }
    if ($grace) {
        $text = "mcp_optional_startup_grace_ms = 5000`n" + $text
        $script:CodexGraceAdded = $true
    }
    [IO.File]::WriteAllText($cfg, $text, (New-Object Text.UTF8Encoding($false)))
    Write-Ok 'set a 30-minute tool time limit, no per-call approval prompts, and a 5 s start-up wait'
}

function Set-OpenCode {
    $oc = Find-Exe 'opencode'
    $dir = Join-Path $HomeDir '.config\opencode'
    if (-not $oc -and -not (Test-Path -LiteralPath $dir)) { Write-Warn2 'OpenCode not found; skipped'; return $false }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $jsonc = Join-Path $dir 'opencode.jsonc'
    $json = Join-Path $dir 'opencode.json'
    $entry = [pscustomobject]@{
        type    = 'local'
        command = @($script:Py, $script:Server)
        enabled = $true
        timeout = 1800000
    }
    # OpenCode merges opencode.json and opencode.jsonc, and creates opencode.jsonc (plain JSON) on its first
    # launch. Edit the first of them that reads as plain JSON. A file with comments is never rewritten;
    # opencode.json is used instead (OpenCode merges it).
    $target = $null
    $cfg = $null
    foreach ($f in @($jsonc, $json)) {
        if (-not (Test-Path -LiteralPath $f)) { continue }
        $raw = [IO.File]::ReadAllText($f)
        if (-not $raw.Trim()) { $target = $f; $cfg = [pscustomobject]@{ '$schema' = 'https://opencode.ai/config.json' }; break }
        $parsed = $null
        try { $parsed = $raw | ConvertFrom-Json } catch { $parsed = $null }
        if ($null -ne $parsed) { $target = $f; $cfg = $parsed; break }
    }
    if ($null -eq $target) {
        if (Test-Path -LiteralPath $json) {
            Write-Warn2 "$json has comments, so it was not edited. Add this under ""mcp"" yourself:"
            Write-Host (@{ $ServerName = $entry } | ConvertTo-Json -Depth 5)
            return $false
        }
        $target = $json
        $cfg = [pscustomobject]@{ '$schema' = 'https://opencode.ai/config.json' }
    } else {
        Backup-Item $target ('opencode-' + (Split-Path $target -Leaf)) | Out-Null
    }
    if (-not ($cfg.PSObject.Properties.Name -contains 'mcp')) {
        $cfg | Add-Member -NotePropertyName 'mcp' -NotePropertyValue ([pscustomobject]@{})
    }
    if ($cfg.mcp.PSObject.Properties.Name -contains $ServerName) {
        $cfg.mcp.PSObject.Properties.Remove($ServerName)
    }
    $cfg.mcp | Add-Member -NotePropertyName $ServerName -NotePropertyValue $entry
    [IO.File]::WriteAllText($target, ($cfg | ConvertTo-Json -Depth 50), (New-Object Text.UTF8Encoding($false)))
    Write-Ok "MCP server '$ServerName' added to $target"
    if (-not $NoSkill) {
        # OpenCode also reads ~/.claude/skills and ~/.agents/skills: a second copy would only cause a warning.
        $shared = @()
        foreach ($name in Get-SkillNames) {
            foreach ($root in @((Join-Path $HomeDir '.claude\skills'), (Join-Path $HomeDir '.agents\skills'))) {
                if (Test-Path -LiteralPath (Join-Path (Join-Path $root $name) 'SKILL.md')) { $shared += $name; break }
            }
        }
        if ($shared.Count -eq (Get-SkillNames).Count) {
            Write-Ok 'skill already visible to OpenCode (it also reads ~/.claude/skills and ~/.agents/skills)'
        } else {
            # ~/.agents/skills is shared with Codex, so setting up Codex later adds no second copy.
            Install-Skills (Join-Path $HomeDir '.agents\skills') 'opencode'
        }
    }
    return $true
}

function Main {
    $ErrorActionPreference = 'Stop'
    Write-Host ''
    Write-Host '  PowerPoint Accessibility MCP installer' -ForegroundColor Cyan
    Write-Host '  https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp'

    Write-Step 'Checking this computer'
    if ([Environment]::OSVersion.Platform -ne 'Win32NT') { throw 'This tool needs Windows: it drives desktop PowerPoint.' }
    if (-not (Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application')) {
        throw 'Desktop PowerPoint was not found. Install Microsoft PowerPoint (Microsoft 365 or Office) and run this again.'
    }
    Write-Ok 'Windows and desktop PowerPoint found'

    $script:App = Join-Path $InstallDir 'app'
    $script:Py = Join-Path $script:App '.venv\Scripts\python.exe'
    $script:Server = Join-Path $script:App 'server.py'

    Write-Step 'Getting uv (provides a private Python)'
    $uv = Get-Uv
    Write-Ok "uv: $uv"

    Write-Step "Getting the code into $($script:App)"
    Get-Code

    Write-Step 'Creating the Python environment (first run downloads Python, about 1 minute)'
    # Keep uv's Python and download cache inside the install folder (see -InstallDir).
    $oldPy = $env:UV_PYTHON_INSTALL_DIR
    $oldCache = $env:UV_CACHE_DIR
    $env:UV_PYTHON_INSTALL_DIR = Join-Path $InstallDir 'python'
    $env:UV_CACHE_DIR = Join-Path $InstallDir 'cache'
    try {
        $venv = Join-Path $script:App '.venv'
        $attempts = @(
            @('--python', '3.12', '--python-preference', 'only-managed'),
            @('--python', '3.12'),
            @('--python', '3.13'),
            @('--python', '3.11'),
            @('--python', '3.10')
        )
        $made = $false
        $errors = @()
        foreach ($a in $attempts) {
            $r = Invoke-Native $uv (@('venv', $venv, '--quiet', '--clear') + $a)
            if ($r.Code -eq 0 -and (Test-Path -LiteralPath $script:Py)) { $made = $true; break }
            $errors += "uv venv $($a -join ' '): $($r.Output)"
        }
        if (-not $made) { throw "creating the environment failed:`n$($errors -join "`n")" }
        $r = Invoke-Native $uv @('pip', 'install', '--python', $script:Py, '--quiet', '-r', (Join-Path $script:App 'requirements.txt'))
        if ($r.Code -ne 0) { throw "installing the Python packages failed:`n$($r.Output)" }
    } finally {
        $env:UV_PYTHON_INSTALL_DIR = $oldPy
        $env:UV_CACHE_DIR = $oldCache
        Remove-Item -LiteralPath (Join-Path $InstallDir 'cache') -Recurse -Force -ErrorAction SilentlyContinue
    }
    $ver = (Invoke-Native $script:Py @('-c', 'import sys; print(sys.version.split()[0])')).Output
    Write-Ok "packages installed (Python $ver)"

    Write-Step 'Checking the server'
    $r = Invoke-Native $script:Py @((Join-Path $script:App 'cli.py'), 'list')
    $tools = @(($r.Output -split "`n") | Where-Object { $_ -match '^[a-z_]+\(' }).Count
    if ($r.Code -ne 0 -or $tools -lt $MinTools) { throw "the server did not start correctly ($tools tools):`n$($r.Output)" }
    Write-Ok "server works: $tools tools"

    $want = @($Apps | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ })
    $auto = $want -contains 'auto'
    $done = @()
    if ($auto -or $want -contains 'claude') {
        Write-Step 'Claude Code'
        if (Set-Claude) { $done += 'Claude Code' }
    }
    if ($auto -or $want -contains 'codex') {
        Write-Step 'Codex'
        if (Set-Codex) { $done += 'Codex' }
    }
    if ($auto -or $want -contains 'opencode') {
        Write-Step 'OpenCode'
        if (Set-OpenCode) { $done += 'OpenCode' }
    }

    $record = [pscustomobject]@{ installed = (Get-Date).ToString('s'); app = $script:App; python = $script:Py; server = $script:Server; server_name = $ServerName; apps = $done; ref = $Ref; codex_grace_added = [bool]$script:CodexGraceAdded }
    $old = Join-Path $InstallDir 'install.json'
    if (Test-Path -LiteralPath $old) {
        # keep the record of a setting an earlier run added
        try { if ((Get-Content -Raw -LiteralPath $old | ConvertFrom-Json).codex_grace_added) { $record.codex_grace_added = $true } } catch { }
    }
    [IO.File]::WriteAllText((Join-Path $InstallDir 'install.json'), ($record | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))

    Write-Host ''
    Write-Host '  Done.' -ForegroundColor Green
    if ($done.Count -gt 0) {
        Write-Host "  Set up for: $($done -join ', ')"
        Write-Host '  Restart the app (or start a new session), then ask, for example:'
    } elseif ($want -contains 'none') {
        Write-Host '  Installed without setting up an app. Register this command in your MCP app:'
        Write-Host "    ""$($script:Py)"" ""$($script:Server)"""
        Write-Host '  Then ask, for example:'
    } else {
        Write-Host '  No app was set up. Install Claude Code, Codex or OpenCode and run this again,'
        Write-Host '  or see docs/manual-setup.md. The server command is:'
        Write-Host "    ""$($script:Py)"" ""$($script:Server)"""
        Write-Host '  Then ask, for example:'
    }
    Write-Host ''
    Write-Host '    Make "C:\path\to\Lecture 05.pptx" WCAG 2.1 AA compliant and save the' -ForegroundColor White
    Write-Host '    _accessible copy in the same folder.' -ForegroundColor White
    Write-Host ''
}

try {
    Main
} catch {
    Write-Host ''
    Write-Host "  Installation stopped: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host '  Help: https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/blob/main/docs/troubleshooting.md'
    Write-Host ''
}
