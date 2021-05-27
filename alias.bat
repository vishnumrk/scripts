@echo off
echo.
set CURRENT_DIR=%~dp0
echo %CURRENT_DIR%

for /f "tokens=1,* delims= " %%a in ("%*") do set ALL_BUT_FIRST=%%b


echo @echo off > %CURRENT_DIR%\%1.bat
echo echo. >> %CURRENT_DIR%\%1.bat
echo %ALL_BUT_FIRST% %%* >> %CURRENT_DIR%\%1.bat
echo Created alias for %1=%ALL_BUT_FIRST%

