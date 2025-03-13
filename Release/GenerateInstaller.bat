echo on
if not "%1"=="" cd %1
del .\GRIPSDirectPrintClientInstaller.7z
del .\GRIPSDirectPrintClientInstaller.exe
del "%temp%\InstallFiles\*.*"
xcopy /s /i ..\*.* "%temp%\InstallFiles"
rem md "%temp%\InstallFiles\Resources"
rem xcopy ..\bin\Release\Resources\*.psd1 "%temp%\InstallFiles\Resources" /S
rem xcopy ..\bin\Release\Certificate.pfx "%temp%\InstallFiles" 
rem pause
del %temp%\GRIPSDirectPrintClientInstaller.7z
.\lzma2200\bin\7zr.exe a -r %temp%\GRIPSDirectPrintClientInstaller.7z "%temp%\InstallFiles\*.*"
copy %temp%\GRIPSDirectPrintClientInstaller.7z
copy /b 7zSD.sfx + .\GRIPSDirectPrintClientInstaller.config.txt + .\GRIPSDirectPrintClientInstaller.7z GRIPSDirectPrintClientInstaller.exe
rmdir %temp%\InstallFiles /S /Q
del %temp%\GRIPSDirectPrintClientInstaller.7z
..\..\grips-signtool\signtool.exe sign /f ..\..\grips-signtool\GRIPS_Codesign.pfx /p "GRIPSCodeSign23~" /tr http://timestamp.sectigo.com /td SHA256 /fd SHA256 GRIPSDirectPrintClientInstaller.exe
del .\GRIPSDirectPrintClientInstaller.7z
rem exit code 0