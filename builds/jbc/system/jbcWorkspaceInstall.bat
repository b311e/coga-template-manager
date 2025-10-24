@echo off
setlocal EnableExtensions

REM === Define agency (HOU, SEN, JBC, LCS, OLLS, OSA) ===
set "AGENCY=JBC"

REM === Map agency names to folder names ===
set "AGENCY_FOLDER=%AGENCY%"
if /i "%AGENCY%"=="OLLS" set "AGENCY_FOLDER=LLS"
if /i "%AGENCY%"=="OSA" set "AGENCY_FOLDER=SAO"

REM === Prod location ====
set "PROD=S:\%AGENCY_FOLDER%\TEMPLATES\%AGENCY% Template Install"

REM === Local template locations ===
set "LOCAL=%APPDATA%\Microsoft\Templates"
set "LOCALOFFICE=%LOCALAPPDATA%\Microsoft\Office"

REM === Install custom Fonts ====
del "%LOCAL%\Document Themes\Theme fonts\%AGENCY% Fonts.xml"
xcopy "%PROD%\%AGENCY% Fonts\%AGENCY% Fonts.xml" "%LOCAL%\Document Themes\Theme fonts" /Y

REM === Install custom Colors ====
del "%LOCAL%\Document Themes\Theme Colors\%AGENCY% Colors.xml"
xcopy "%PROD%\%AGENCY% Colors\%AGENCY% Colors.xml" "%LOCAL%\Document Themes\Theme Colors" /Y

REM === Install custom Theme ====
del "%LOCAL%\Document Themes\%AGENCY% Theme.thmx"
xcopy "%PROD%\%AGENCY% Theme\%AGENCY% Theme.thmx" "%LOCAL%\Document Themes" /Y

REM === Install custom Normal ===
del "%LOCAL%\Normal_Backup.dotm"
ren "%LOCAL%\Normal.dotm" "Normal_Backup.dotm"
xcopy "%PROD%\%AGENCY% Normal\Normal.dotm" "%LOCAL%" /Y

REM === Install custom UI ===
del "%LOCALOFFICE%\Word.officeUI"
xcopy "%PROD%\%AGENCY% Office UI\Word.officeUI" "%LOCALOFFICE%" /Y
del "%LOCALOFFICE%\Excel.officeUI"
xcopy "%PROD%\%AGENCY% Office UI\Excel.officeUI" "%LOCALOFFICE%" /Y

REM === Install Excel Startup Templates ===
ren "%APPDATA%\Microsoft\Excel\XLSTART" "XLSTART Backup"
mkdir "%APPDATA%\Microsoft\Excel\XLSTART"
xcopy "%PROD%\%AGENCY% XLSTART\Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" /Y
xcopy "%PROD%\%AGENCY% XLSTART\Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" /Y
xcopy "%PROD%\%AGENCY% XLSTART\PERSONAL.XLSB" "%APPDATA%\Microsoft\Excel\XLSTART" /Y

exit
