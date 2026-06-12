@echo off
setlocal
set "_psf=%TEMP%\templx_unpack_%RANDOM%.ps1"
set "_self=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=Get-Content -LiteralPath $env:_self -Raw; $m='#'+'PSCODE'+'#'; [System.IO.File]::WriteAllText($env:_psf, $s.Substring($s.IndexOf($m)), (New-Object System.Text.UTF8Encoding($false)))"
powershell -NoProfile -ExecutionPolicy Bypass -File "%_psf%" %*
del "%_psf%" 2>nul
endlocal & exit /b

#PSCODE#
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
foreach ($p in $args) {
    $file = (Resolve-Path -LiteralPath $p).Path
    [System.IO.Compression.ZipFile]::ExtractToDirectory($file, $file + "_unpacked")
}
