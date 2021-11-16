@ECHO OFF

:choice1
ECHO Warning! this will delete existing lookout_fix_version-beta-tb.xpi
set /P c=Continue [Y/N]?
if /I "%c%" EQU "Y" goto :old_del
if /I "%c%" EQU "N" goto :exit
goto :choice1

:old_del
del lookout_fix_version-beta-tb.xpi

:choice2
ECHO Warning! this will update the wrapper-apis from Github profile.
set /P c=Continue [Y/N]?
if /I "%c%" EQU "Y" goto :update
if /I "%c%" EQU "N" goto :build
goto :choice2

:update
for /f "tokens=2,*" %%a in ('REG QUERY "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Personal"') DO (SET docsdir=%%b)

copy /Y "%docsdir%\GitHub\addon-developer-support\wrapper-apis\WindowListener\*" "%docsdir%\GitHub\LookOut-fix-version\src\api\WindowListener\"

:build
for /f "tokens=2,*" %%a in ('REG QUERY "HKCU\SOFTWARE\7-Zip" /v "Path"') DO (SET zipdir=%%b)

"%zipdir%\7zG.exe" a -tzip lookout_fix_version-beta-tb.xpi _locales api chrome defaults background.js changes.txt LICENSE manifest.json

:exit

echo Goodbye
pause
exit