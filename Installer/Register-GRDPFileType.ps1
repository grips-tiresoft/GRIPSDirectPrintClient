<#
.SYNOPSIS
    Registers the .grdp file type with Windows for GRIPS Direct Print.

.DESCRIPTION
    This script registers the .grdp file extension with the Windows registry.
    It associates .grdp files with the VBScript handler and sets appropriate
    icons and display names for all users on the system.

.NOTES
    - Must be run with Administrator privileges
    - Registers file type for all users (HKEY_CLASSES_ROOT)
    - VBScript path: C:\ProgramData\GRIPSDirectPrintClient\Print-GRDPFile.vbs
    - Icon path: C:\ProgramData\GRIPSDirectPrintClient\grips.ico

.EXAMPLE
    .\Register-GRDPFileType.ps1
#>

# Requires Administrator privileges
#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$fileExtension = ".grdp"
$progId = "GRIPS.DirectPrint.Archive"
$friendlyTypeName = "GRIPS Direct Print Archive"
$openWithDisplayName = "GRIPS Direct Print"
$vbsScriptPath = "C:\ProgramData\GRIPSDirectPrintClient\Print-GRDPFile.vbs"
$iconPath = "C:\ProgramData\GRIPSDirectPrintClient\grips.ico"

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "GRIPS Direct Print File Type Registration" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# Check if VBScript file exists
if (-not (Test-Path $vbsScriptPath)) {
    Write-Host "WARNING: VBScript file not found at: $vbsScriptPath" -ForegroundColor Yellow
    Write-Host "The file type will be registered, but will not work until this file exists." -ForegroundColor Yellow
    Write-Host ""
}

# Check if icon file exists
if (-not (Test-Path $iconPath)) {
    Write-Host "WARNING: Icon file not found at: $iconPath" -ForegroundColor Yellow
    Write-Host "The file type will be registered, but will use default icon." -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Registering file type with the following settings:" -ForegroundColor White
Write-Host "  File Extension: $fileExtension" -ForegroundColor Gray
Write-Host "  ProgID: $progId" -ForegroundColor Gray
Write-Host "  Friendly Name: $friendlyTypeName" -ForegroundColor Gray
Write-Host "  Open With Display: $openWithDisplayName" -ForegroundColor Gray
Write-Host "  VBScript Handler: $vbsScriptPath" -ForegroundColor Gray
Write-Host "  Icon: $iconPath" -ForegroundColor Gray
Write-Host ""

try {
    # Ensure HKCR: drive is available
    if (-not (Test-Path "HKCR:")) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    }

    # Register the file extension
    Write-Host "1. Registering file extension $fileExtension..." -ForegroundColor Green
    $extKey = "HKCR:\$fileExtension"
    
    if (Test-Path $extKey) {
        Write-Host "   Extension already registered, updating..." -ForegroundColor Yellow
    }
    
    New-Item -Path $extKey -Force | Out-Null
    New-ItemProperty -Path $extKey -Name "(Default)" -Value $progId -Force | Out-Null
    Write-Host "   [OK] Extension registered" -ForegroundColor Green
    Write-Host ""

    # Register the ProgID
    Write-Host "2. Registering ProgID $progId..." -ForegroundColor Green
    $progIdKey = "HKCR:\$progId"
    
    New-Item -Path $progIdKey -Force | Out-Null
    New-ItemProperty -Path $progIdKey -Name "(Default)" -Value $friendlyTypeName -Force | Out-Null
    New-ItemProperty -Path $progIdKey -Name "FriendlyTypeName" -Value $friendlyTypeName -Force | Out-Null
    Write-Host "   [OK] ProgID registered" -ForegroundColor Green
    Write-Host ""

    # Set the default icon
    Write-Host "3. Setting default icon..." -ForegroundColor Green
    $iconKey = "$progIdKey\DefaultIcon"
    New-Item -Path $iconKey -Force | Out-Null
    New-ItemProperty -Path $iconKey -Name "(Default)" -Value $iconPath -Force | Out-Null
    Write-Host "   [OK] Icon set" -ForegroundColor Green
    Write-Host ""

    # Register the shell command to open the file
    Write-Host "4. Registering shell open command..." -ForegroundColor Green
    $shellKey = "$progIdKey\shell"
    $openKey = "$shellKey\open"
    $commandKey = "$openKey\command"
    
    New-Item -Path $shellKey -Force | Out-Null
    New-Item -Path $openKey -Force | Out-Null
    New-ItemProperty -Path $openKey -Name "(Default)" -Value $openWithDisplayName -Force | Out-Null
    New-ItemProperty -Path $openKey -Name "FriendlyAppName" -Value $openWithDisplayName -Force | Out-Null
    
    New-Item -Path $commandKey -Force | Out-Null
    # Use wscript.exe to run the VBScript with the file as an argument
    $commandValue = '"C:\Windows\System32\wscript.exe" "' + $vbsScriptPath + '" "%1"'
    New-ItemProperty -Path $commandKey -Name "(Default)" -Value $commandValue -Force | Out-Null
    Write-Host "   [OK] Shell command registered" -ForegroundColor Green
    Write-Host ""

    # Notify Windows Explorer that file associations have changed
    Write-Host "5. Notifying Windows Explorer of changes..." -ForegroundColor Green
    
    # Define the SHChangeNotify function signature
    $signature = @'
    [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
'@
    
    # Add the type if it doesn't already exist
    try {
        Add-Type -MemberDefinition $signature -Name "Functions" -Namespace "Shell32" -ErrorAction Stop
    } catch {
        # Type already exists, which is fine
        if ($_.Exception.Message -notlike "*already exists*") {
            throw
        }
    }
    
    # SHCNE_ASSOCCHANGED = 0x08000000, SHCNF_IDLIST = 0x0000
    [Shell32.Functions]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)
    
    Write-Host "   [OK] Explorer notified" -ForegroundColor Green
    Write-Host ""

    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Registration completed successfully!" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "The .grdp file type has been registered for all users." -ForegroundColor White
    Write-Host "Files with .grdp extension will now:" -ForegroundColor White
    Write-Host "  - Show as 'GRIPS Direct Print Archive' in Explorer" -ForegroundColor Gray
    Write-Host "  - Display the custom icon (if available)" -ForegroundColor Gray
    Write-Host "  - Open with 'GRIPS Direct Print' when double-clicked" -ForegroundColor Gray
    Write-Host "  - Execute: $vbsScriptPath" -ForegroundColor Gray
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "ERROR: Registration failed!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit..."
    exit 1
}
Read-Host "Press Enter to exit..."
