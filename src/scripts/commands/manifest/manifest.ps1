# Manifest Generator for templx (PowerShell)
# Usage: manifest <command> [options]

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Command,
    
    [Parameter(Position=1)]
    [string]$Agency
)

$ErrorActionPreference = "Stop"

# Get the script directory to find the bash script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BashScript = Join-Path $ScriptDir "manifest"

# Check if Git Bash is available
$GitBash = "C:\Program Files\Git\bin\bash.exe"
if (-not (Test-Path $GitBash)) {
    Write-Host "Error: Git Bash not found at: $GitBash"
    Write-Host "This script requires Git Bash. Please install Git for Windows."
    exit 1
}

# Run the bash script
if ($Agency) {
    & $GitBash -c "bash '$BashScript' '$Command' '$Agency'"
} else {
    & $GitBash -c "bash '$BashScript' '$Command'"
}
