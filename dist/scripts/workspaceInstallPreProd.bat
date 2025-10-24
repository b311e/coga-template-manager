@echo off
setlocal EnableExtensions

REM === Define agency (HOU, SEN, JBC, LCS, OLLS, OSA) ===
echo Available agencies: HOU, SEN, JBC, LCS, OLLS, OSA
set /p "AGENCY=Enter agency: "

REM === Pre-Prod location ===
set "PREPROD=S:\LIS\Accessibility\Assets\Templates\%AGENCY%\%AGENCY% Template Install\PreProd"

REM === Local template locations ===
set "LOCAL=%APPDATA%\Microsoft\Templates"
set "LOCALOFFICE=%LOCALAPPDATA%\Microsoft\Office"

REM === Install agency Fonts ===
del "%LOCAL%\Document Themes\Theme fonts\%AGENCY% Fonts.xml"
xcopy "%PREPROD%\%AGENCY% Fonts\%AGENCY% Fonts.xml" "%LOCAL%\Document Themes\Theme fonts" /Y

REM === Install agency Colors ===
del "%LOCAL%\Document Themes\Theme Colors\%AGENCY% Colors.xml"
xcopy "%PREPROD%\%AGENCY% Colors\%AGENCY% Colors.xml" "%LOCAL%\Document Themes\Theme Colors" /Y

REM === Install agency Theme ===
del "%LOCAL%\Document Themes\%AGENCY% Theme.thmx"
xcopy "%PREPROD%\%AGENCY% Theme\%AGENCY% Theme.thmx" "%LOCAL%\Document Themes" /Y

REM === Install agency Normal ===
ren "%LOCAL%\Normal.dotm" "Normal_Backup.dotm"
xcopy "%PREPROD%\%AGENCY% Normal\Normal.dotm" "%LOCAL%" /Y

REM === Install agency UI ===
del "%LOCALOFFICE%\Word.officeUI"
xcopy "%PREPROD%\%AGENCY% Office UI\Word.officeUI" "%LOCALOFFICE%" /Y
del "%LOCALOFFICE%\Excel.officeUI"
xcopy "%PREPROD%\%AGENCY% Office UI\Excel.officeUI" "%LOCALOFFICE%" /Y

REM === Install agency Excel startup Templates ===
ren "%APPDATA%\Microsoft\Excel\XLSTART" "XLSTART Backup"
mkdir "%APPDATA%\Microsoft\Excel\XLSTART"
xcopy "%PREPROD%\%AGENCY% XLSTART\Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" /Y
xcopy "%PREPROD%\%AGENCY% XLSTART\Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" /Y
xcopy "%PREPROD%\%AGENCY% XLSTART\PERSONAL.XLSB" "%APPDATA%\Microsoft\Excel\XLSTART" /Y

exit
