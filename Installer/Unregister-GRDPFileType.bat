@echo off
REM Unregister GRDP File Type
REM This batch file runs the PowerShell unregistration script with Administrator privileges

echo.
echo ===============================================
echo  GRIPS Direct Print File Type Unregistration
echo ===============================================
echo.
echo This will unregister .grdp files from Windows.
echo Administrator privileges are required.
echo.
pause

REM Run PowerShell script as Administrator
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0Unregister-GRDPFileType.ps1""' -Verb RunAs}"

echo.
echo Unregistration script executed.
echo Check the PowerShell window for results.
echo.
pause
