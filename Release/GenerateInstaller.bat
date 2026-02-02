echo on
if not "%1"=="" cd %1
del .\GRIPSDirectPrintClientInstaller.7z
del .\GRIPSDirectPrintClientInstaller.exe
del "%temp%\InstallFiles\*.*"
xcopy /s /i ..\*.* "%temp%\InstallFiles"

del %temp%\GRIPSDirectPrintClientInstaller.7z
.\lzma2200\bin\7zr.exe a -r %temp%\GRIPSDirectPrintClientInstaller.7z "%temp%\InstallFiles\*.*"
copy %temp%\GRIPSDirectPrintClientInstaller.7z .\GRIPSDirectPrintClientInstaller.7z
copy /b 7zSD.sfx + .\GRIPSDirectPrintClientInstaller.config.txt + .\GRIPSDirectPrintClientInstaller.7z GRIPSDirectPrintClientInstaller.exe
rmdir %temp%\InstallFiles /S /Q
del %temp%\GRIPSDirectPrintClientInstaller.7z

set signToolPath=%~dp0..\..\grips-signtool\AzureArtifactSigning\microsoft.windows.sdk.buildtools\10.0.26100.7463\bin\10.0.26100.0\x64\signtool.exe
set DlibDllPath=%~dp0..\..\grips-signtool\AzureArtifactSigning\microsoft.artifactsigning.client\1.0.115\bin\x64\Azure.CodeSigning.Dlib.dll
set metaDataFilePath=%~dp0..\..\grips-signtool\AzureArtifactSigning\GRIPS-codesign.json

echo Signing installer...
echo %signToolPath%
echo %DlibDllPath%
echo %metaDataFilePath%

"%signToolPath%" sign /v /debug /fd SHA256 /tr "http://timestamp.acs.microsoft.com" /td SHA256 /dlib "%DlibDllPath%" /dmdf "%metaDataFilePath%" GRIPSDirectPrintClientInstaller.exe

del .\GRIPSDirectPrintClientInstaller.7z

echo Installer generation complete.