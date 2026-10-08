<#
.SYNOPSIS
    Removes PowerPoint Accessibility MCP (server "ppt-a11y"): its registration in Claude Code, Codex and
    OpenCode, the skill copies the installer made, and the install folder. Decks are never touched.

.EXAMPLE
    irm https://raw.githubusercontent.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/main/uninstall.ps1 | iex
#>
param(
    [string]$InstallDir = (Join-Path $env:USERPROFILE '.ppt-a11y'),
    [string]$ServerName = 'ppt-a11y',
    # Keep the install folder (and its backups) and only remove the registrations and skills.
    [switch]$KeepFiles,
    [string]$HomeDir = $env:USERPROFILE
)

function Write-Ok([string]$Text) { Write-Host "    [ok] $Text" -ForegroundColor Green }
function Write-Note([string]$Text) { Write-Host "    $Text" }

function Invoke-Native {
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

function Remove-InstalledSkills([string]$Root) {
    # Only folders the installer created (they carry a marker file) are removed.
    if (-not (Test-Path -LiteralPath $Root)) { return }
    foreach ($d in Get-ChildItem -LiteralPath $Root -Directory) {
        if (Test-Path -LiteralPath (Join-Path $d.FullName '.installed-by-ppt-a11y')) {
            Remove-Item -LiteralPath $d.FullName -Recurse -Force
            Write-Ok "removed skill $($d.FullName)"
        }
    }
}

function Main {
    $ErrorActionPreference = 'Stop'
    Write-Host ''
    Write-Host '  PowerPoint Accessibility MCP uninstaller' -ForegroundColor Cyan

    $claude = Find-Exe 'claude'
    if ($claude) {
        $r = Invoke-Native $claude @('mcp', 'remove', $ServerName, '--scope', 'user')
        if ($r.Code -eq 0) { Write-Ok "Claude Code: removed '$ServerName'" } else { Write-Note "Claude Code: '$ServerName' was not registered" }
    }
    $codex = Find-Exe 'codex'
    if ($codex) {
        $r = Invoke-Native $codex @('mcp', 'remove', $ServerName)
        if ($r.Code -eq 0) { Write-Ok "Codex: removed '$ServerName'" } else { Write-Note "Codex: '$ServerName' was not registered" }
    }
    # The installer may have added one top-level Codex setting (a 5 s start-up wait); remove only that line.
    $record = $null
    try { $record = Get-Content -Raw -LiteralPath (Join-Path $InstallDir 'install.json') | ConvertFrom-Json } catch { }
    $codexHome = Join-Path $HomeDir '.codex'
    if ($env:CODEX_HOME) { $codexHome = $env:CODEX_HOME }
    $toml = Join-Path $codexHome 'config.toml'
    if ($record -and $record.codex_grace_added -and (Test-Path -LiteralPath $toml)) {
        $text = [IO.File]::ReadAllText($toml)
        $new = [regex]::Replace($text, '(?m)^mcp_optional_startup_grace_ms = 5000\r?\n', '', 1)
        if ($new -ne $text) {
            [IO.File]::WriteAllText($toml, $new, (New-Object Text.UTF8Encoding($false)))
            Write-Ok 'Codex: removed the start-up wait setting the installer added'
        }
    }
    foreach ($name in @('opencode.json', 'opencode.jsonc')) {
        $json = Join-Path $HomeDir ".config\opencode\$name"
        if (-not (Test-Path -LiteralPath $json)) { continue }
        $raw = [IO.File]::ReadAllText($json)
        $cfg = $null
        try { $cfg = $raw | ConvertFrom-Json } catch { $cfg = $null }
        if ($cfg -and $cfg.PSObject.Properties.Name -contains 'mcp' -and $cfg.mcp.PSObject.Properties.Name -contains $ServerName) {
            $cfg.mcp.PSObject.Properties.Remove($ServerName)
            [IO.File]::WriteAllText($json, ($cfg | ConvertTo-Json -Depth 50), (New-Object Text.UTF8Encoding($false)))
            Write-Ok "OpenCode: removed '$ServerName' from $json"
        } elseif (-not $cfg -and $raw -match [regex]::Escape("""$ServerName""")) {
            Write-Note "OpenCode: $json has comments; remove ""$ServerName"" under ""mcp"" yourself"
        }
    }

    foreach ($root in @('.claude\skills', '.agents\skills', '.config\opencode\skills', '.codex\skills')) {
        Remove-InstalledSkills (Join-Path $HomeDir $root)
    }

    if (-not $KeepFiles -and (Test-Path -LiteralPath $InstallDir)) {
        try {
            Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction Stop
            Write-Ok "removed $InstallDir"
        } catch {
            Write-Note "could not remove $InstallDir (is an app still running the server?). Close the apps and delete the folder."
        }
    }
    Write-Host ''
    Write-Host '  Done. Restart your AI apps. Your decks and their _accessible copies were not touched.' -ForegroundColor Green
    Write-Host ''
}

try { Main } catch { Write-Host "  Uninstall stopped: $($_.Exception.Message)" -ForegroundColor Red }
