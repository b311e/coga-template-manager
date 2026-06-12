@echo off
setlocal
set "_psf=%TEMP%\templx_pack_%RANDOM%.ps1"
set "_self=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=Get-Content -LiteralPath $env:_self -Raw; $m='#'+'PSCODE'+'#'; [System.IO.File]::WriteAllText($env:_psf, $s.Substring($s.IndexOf($m)), (New-Object System.Text.UTF8Encoding($false)))"
powershell -NoProfile -ExecutionPolicy Bypass -File "%_psf%" %*
del "%_psf%" 2>nul
endlocal & exit /b

#PSCODE#

Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
foreach ($p in $args) {
    $folder = (Resolve-Path -LiteralPath $p).Path.TrimEnd('\')

    # File type comes from [Content_Types].xml. Keep 'themeManager' before 'pptx'.

    $ct  = Get-Content -LiteralPath (Join-Path $folder '[Content_Types].xml') -Raw -ErrorAction SilentlyContinue
    $ext = switch -Regex ($ct) {
        'officedocument\.themeManager'       { 'thmx'; break }
        'wordprocessingml\.document\.main'   { 'docx'; break }
        'wordprocessingml\.template\.main'   { 'dotx'; break }
        'ms-word\.document\.macroEnabled'    { 'docm'; break }
        'ms-word\.template\.macroEnabled'    { 'dotm'; break }
        'spreadsheetml\.sheet\.main'         { 'xlsx'; break }
        'spreadsheetml\.template\.main'      { 'xltx'; break }
        'ms-excel\.sheet\.macroEnabled'      { 'xlsm'; break }
        'ms-excel\.template\.macroEnabled'   { 'xltm'; break }
        'presentationml\.presentation\.main' { 'pptx'; break }
        'presentationml\.template\.main'     { 'potx'; break }
        'presentationml\.slideshow\.main'    { 'ppsx'; break }
        'ms-powerpoint\.presentation\.macro' { 'pptm'; break }
        'ms-powerpoint\.template\.macro'     { 'potm'; break }
        'ms-powerpoint\.slideshow\.macro'    { 'ppsm'; break }
    }

    if (-not $ext) { Write-Host "Skipping (not a recognized Office package): $folder"; continue }

    $base = [System.IO.Path]::GetFileNameWithoutExtension(($folder -replace '_unpacked$', ''))
    $out  = Join-Path ([System.IO.Path]::GetDirectoryName($folder)) ($base + '_packed.' + $ext)
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

    # Build the zip with forward slashes not backward slashes

    $zip = [System.IO.Compression.ZipFile]::Open($out, 'Create')
    $files = Get-ChildItem -LiteralPath $folder -Recurse -File |
        Sort-Object @{ Expression = { if ($_.Name -eq '[Content_Types].xml') { 0 } else { 1 } } }, FullName
    foreach ($f in $files) {
        $name  = $f.FullName.Substring($folder.Length).TrimStart('\').Replace('\', '/')
        $entry = $zip.CreateEntry($name)
        $es = $entry.Open(); $in = [System.IO.File]::OpenRead($f.FullName)
        $in.CopyTo($es); $in.Dispose(); $es.Dispose()
    }
    $zip.Dispose()
}
