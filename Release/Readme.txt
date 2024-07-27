Run GenerateInstaller.bat to create new self-extracting EXE

Before running, clone the grips-signtool repo from https://github.com/goodyear/grips-signtool.git
This should be cloned to same parent folder as this one.
If this is not done then code-signing will fail.

Remember to set compatibility settings for Windows 7 after to avoid "Did this program run correctly?" message.
Copy GRIPSDirectPrintClientInstaller.exe to target machine and run as administrator (should prompt for admin permissions if run normally)

Once the installer EXE is created, create new release on https://github.com/grips-tiresoft/GRIPSDirectPrintClient/releases and upload the installer EXE.

To fully uninstall:
# TODO: Run the '%programdata%\GRIPSDirectPrintClient\Installer\Uninstall.ps1' script in powershell
