@echo off
REM Register GRDP File Type
REM This batch file runs the PowerShell registration script with Administrator privileges

echo.
echo ===============================================
echo  GRIPS Direct Print File Type Registration
echo ===============================================
echo.
echo This will register .grdp files with Windows.
echo Administrator privileges are required.
echo.
pause

REM Run PowerShell script as Administrator
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0Register-GRDPFileType.ps1""' -Verb RunAs}"

echo.
echo Registration script executed.
echo Check the PowerShell window for results.
echo.
pause
