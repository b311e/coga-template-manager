# WordprocessingML Automated Cleanup Script (PowerShell)
# Performs automated cleanup tasks from the cleanup checklist
# Usage: cleanup wordml <expanded-template-dir>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$TemplateDir
)

$ErrorActionPreference = "Stop"

if (-not $TemplateDir) {
    Write-Host "Usage: cleanup wordml <expanded-template-dir>"
    Write-Host ""
    Write-Host "Example:"
    Write-Host "  cleanup wordml C:\path\to\template_expanded"
    Write-Host ""
    Write-Host "This script performs automated cleanup on unpacked WordprocessingML files:"
    Write-Host "  - Removes RSID attributes and blocks"
    Write-Host "  - Cleans up Normal style"
    Write-Host "  - Removes noProof, iCs, szCs, bCs elements"
    Write-Host "  - Removes bidirectional language settings"
    Write-Host "  - Cleans up empty rPr and pPr blocks"
    exit 1
}

# Validate directory exists
if (-not (Test-Path $TemplateDir -PathType Container)) {
    Write-Host "Error: Directory not found: $TemplateDir"
    exit 1
}

# Get the script directory to find the bash script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BashScript = Join-Path $ScriptDir "cleanup"

# Check if Git Bash is available
$GitBash = "C:\Program Files\Git\bin\bash.exe"
if (-not (Test-Path $GitBash)) {
    Write-Host "Error: Git Bash not found at: $GitBash"
    Write-Host "This script requires Git Bash. Please install Git for Windows."
    exit 1
}

# Convert Windows path to Unix-style path for bash
$UnixPath = $TemplateDir -replace '\\', '/' -replace '^([A-Z]):', '/$1'

# Run the bash script
& $GitBash -c "bash '$BashScript' '$UnixPath'"
