Run GenerateInstaller.bat to create new self-extracting EXE

Before running, clone the grips-signtool repo from https://github.com/goodyear/grips-signtool.git
This should be cloned to same parent folder as this one.
If this is not done then code-signing will fail.

If you change anything, always increase Script version at the start of Run-GRIPSDirectPrintProcessor.ps1 and make sure the script is resigned:
.\CreateSignedScript.ps1 -Path .\Run-GRIPSDirectPrintProcessor.ps1

Copy GRIPSDirectPrintClientInstaller.exe to target machine and run as administrator (should prompt for admin permissions if run normally)

Once the installer EXE is created, create new release on https://github.com/grips-tiresoft/GRIPSDirectPrintClient/releases and upload the installer EXE.

To fully uninstall:
# TODO: Run the '%programdata%\GRIPSDirectPrintClient\Installer\Uninstall.ps1' script in powershell
# e.g. ..\Installer\nssm-2.24\win64\nssm.exe uninstall "GRIPSDirectPrint Client Service"
