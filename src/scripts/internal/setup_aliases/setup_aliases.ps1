# Setup script to add project commands to PowerShell session
# Run this with: . .\src\scripts\setup_aliases\setup_aliases.ps1
# Or add to your $PROFILE for permanent aliases

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptsPath = Split-Path -Parent $ScriptDir
$ProjectRoot = Resolve-Path (Join-Path $ScriptsPath "..\..")

Write-Host "Setting up templx PowerShell aliases..."
Write-Host ""

# Define all command functions
function global:create { & "$ScriptsPath\create\create.ps1" @args }
function global:unpack { & "$ScriptsPath\unpack\unpack.ps1" @args }
function global:pack { & "$ScriptsPath\pack\pack.ps1" @args }
function global:replace-styles { & "$ScriptsPath\replace_styles\replace-styles.ps1" @args }
function global:replace-styles-from-snippet { & "$ScriptsPath\replace_styles\replace-styles-from-snippet.ps1" @args }
function global:cleanup { & "$ScriptsPath\commands\cleanup\cleanup.ps1" @args }
function global:list-styles { & "$ScriptsPath\list-styles.ps1" @args }
function global:manifest { & "$ScriptsPath\commands\manifest\manifest.ps1" @args }
function global:validate { & "$ScriptsPath\validate\validate.ps1" @args }

Write-Host "PowerShell aliases configured for current session!"
Write-Host ""
Write-Host "Available commands:"
Write-Host "  create - Create new OpenXML templates/documents"
Write-Host "  unpack - Unpack OpenXML file to expanded folder"
Write-Host "  pack - Pack directory to OpenXML file"
Write-Host "  replace-styles - Copy styles from source to target Word doc"
Write-Host "  replace-styles-from-snippet - Replace styles from XML snippet file"
Write-Host "  cleanup wordml - Automated cleanup of WordprocessingML files"
Write-Host "  list-styles - Generate style list for template"
Write-Host "  manifest - Manage template manifests and registry"
Write-Host "  validate - Validate OOXML files and generate reports"
Write-Host ""
Write-Host "Examples:"
Write-Host "  create xl-template Book"
Write-Host "  validate core\templates\normal\src\Normal.dotm"
Write-Host "  list-styles builds\jbc\workspace\jbcNormal"
