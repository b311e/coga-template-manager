# Setup script to expose templx commands in the current PowerShell session.
# Run with:    . .\src\scripts\internal\setup_aliases\setup_aliases.ps1
# Or add to your $PROFILE to load every session.
#
# The bash scripts under src/scripts/bin/ are the canonical implementations.
# This file defines thin PowerShell wrappers that invoke them through Git Bash.

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BinPath = Resolve-Path (Join-Path $ScriptDir "..\..\bin") | Select-Object -ExpandProperty Path

$GitBash = "C:\Program Files\Git\bin\bash.exe"
if (-not (Test-Path $GitBash)) {
    Write-Host "Error: Git Bash not found at: $GitBash"
    Write-Host "This script requires Git for Windows. Install it from https://gitforwindows.org/"
    return
}

# Convert a Windows path to a Git Bash path (C:\foo -> /c/foo)
function script:ConvertTo-BashPath($path) {
    $p = $path -replace '\\', '/'
    if ($p -match '^([A-Za-z]):/(.*)') {
        return "/$($matches[1].ToLower())/$($matches[2])"
    }
    return $p
}

function script:Invoke-TemplxCommand($name) {
    $bashPath = ConvertTo-BashPath (Join-Path $BinPath $name)
    & $GitBash $bashPath @args
}

function global:templx    { Invoke-TemplxCommand 'templx' @args }
function global:pack      { Invoke-TemplxCommand 'pack' @args }
function global:unpack    { Invoke-TemplxCommand 'unpack' @args }
function global:create    { Invoke-TemplxCommand 'create' @args }
function global:validate  { Invoke-TemplxCommand 'validate' @args }
function global:xpathsel  { Invoke-TemplxCommand 'xpathsel' @args }
function global:style     { Invoke-TemplxCommand 'style' @args }
function global:inventory { Invoke-TemplxCommand 'inventory' @args }
function global:cleanup   { Invoke-TemplxCommand 'cleanup' @args }
function global:manifest  { Invoke-TemplxCommand 'manifest' @args }

Write-Host "templx PowerShell aliases configured for current session."
Write-Host ""
Write-Host "Top-level commands available:"
Write-Host "  templx pack unpack create validate xpathsel"
Write-Host "  style inventory cleanup manifest"
Write-Host ""
Write-Host "Run 'templx help' for the full command list."
