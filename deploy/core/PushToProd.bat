@echo off
setlocal EnableExtensions

REM === Define agency (HOU, SEN, JBC, LCS, OLLS, OSA) ===
echo Available agencies: HOU, SEN, JBC, LCS, OLLS, OSA
set /p "AGENCY=Enter agency: "

REM === Map agency names to folder names ===
set "AGENCY_FOLDER=%AGENCY%"
if /i "%AGENCY%"=="OLLS" set "AGENCY_FOLDER=LLS"
if /i "%AGENCY%"=="OSA" set "AGENCY_FOLDER=SAO"

REM === PreProd and Prod locations ===
set "PREPROD=S:\LIS\Accessibility\Assets\Templates\%AGENCY%\%AGENCY% Template Install\PreProd"
set "PROD=S:\%AGENCY_FOLDER%\TEMPLATES\%AGENCY% Template Install"

REM === PreProd subfolders ===
set "PP_COLORS=%PREPROD%\%AGENCY% Colors"
set "PP_FONTS=%PREPROD%\%AGENCY% Fonts"
set "PP_INSTALL=%PREPROD%\%AGENCY% Install Scripts"
set "PP_NORMAL=%PREPROD%\%AGENCY% Normal"
set "PP_OFFICEUI=%PREPROD%\%AGENCY% Office UI"
set "PP_THEME=%PREPROD%\%AGENCY% Theme"
set "PP_XLSTART=%PREPROD%\%AGENCY% XLSTART"

REM === Clean and recreate destination subfolders in Prod ===
rmdir /s /q "%PROD%" 2>nul
mkdir "%PROD%" 2>nul
mkdir "%PROD%\%AGENCY% Colors" 2>nul
mkdir "%PROD%\%AGENCY% Fonts" 2>nul
mkdir "%PROD%\%AGENCY% Normal" 2>nul
mkdir "%PROD%\%AGENCY% Office UI" 2>nul
mkdir "%PROD%\%AGENCY% Theme" 2>nul
mkdir "%PROD%\%AGENCY% XLSTART" 2>nul

REM === Ensure folders remain hidden ===
attrib +h "%PROD%" 2>nul
attrib +h "%PROD%\%AGENCY% Colors" 2>nul
attrib +h "%PROD%\%AGENCY% Fonts" 2>nul
attrib +h "%PROD%\%AGENCY% Normal" 2>nul
attrib +h "%PROD%\%AGENCY% Office UI" 2>nul
attrib +h "%PROD%\%AGENCY% Theme" 2>nul
attrib +h "%PROD%\%AGENCY% XLSTART" 2>nul

REM === Copy PreProd files to Prod ===
xcopy "%PP_COLORS%\%AGENCY% Colors.xml" "%PROD%\%AGENCY% Colors\" /Y
xcopy "%PP_FONTS%\%AGENCY% Fonts.xml" "%PROD%\%AGENCY% Fonts\" /Y
xcopy "%PP_INSTALL%\%AGENCY%TemplateInstall.bat" "%PROD%\" /Y
xcopy "%PP_NORMAL%\Normal.dotm" "%PROD%\%AGENCY% Normal\" /Y
xcopy "%PP_OFFICEUI%\Word.officeUI" "%PROD%\%AGENCY% Office UI\" /Y
xcopy "%PP_OFFICEUI%\Excel.officeUI" "%PROD%\%AGENCY% Office UI\" /Y
xcopy "%PP_THEME%\%AGENCY% Theme.thmx" "%PROD%\%AGENCY% Theme\" /Y
xcopy "%PP_XLSTART%\Book.xltx" "%PROD%\%AGENCY% XLSTART\" /Y
xcopy "%PP_XLSTART%\Sheet.xltx" "%PROD%\%AGENCY% XLSTART\" /Y
xcopy "%PP_XLSTART%\PERSONAL.XLSB" "%PROD%\%AGENCY% XLSTART\" /Y

exit
