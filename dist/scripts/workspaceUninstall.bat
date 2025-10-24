@echo off
setlocal

REM === Set the folder paths here ===
set "FONTS=%APPDATA%\Microsoft\Templates\Document Themes\Theme fonts"
set "COLORS=%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors"
set "THEME=%APPDATA%\Microsoft\Templates\Document Themes"
set "OFFICEUI=%LOCALAPPDATA%\Microsoft\Office"
set "NORMAL=%APPDATA%\Microsoft\Templates"

echo Deleting all custom Normal, themes, colors, and fonts...

if exist "%FONTS%\*.xml" del "%FONTS%\*.xml"
if exist "%COLORS%\*.xml" del "%COLORS%\*.xml"
if exist "%THEME%\*.thmx" del "%THEME%\*.thmx"
if exist "%OFFICEUI%\*.officeUI" del "%OFFICEUI%\*.officeUI"

if exist "%NORMAL%\Normal_Backup.dotm" del "%NORMAL%\Normal_Backup.dotm
if exist "%NORMAL%\Normal.dotm" ren "%NORMAL%\Normal.dotm" Normal_Backup.dotm

echo Done.
endlocal

exit