<#
.SYNOPSIS
    Unregisters the .grdp file type from Windows.

.DESCRIPTION
    This script removes the .grdp file extension registration from the Windows registry.
    It removes all registry keys associated with the GRIPS Direct Print file type.

.NOTES
    - Must be run with Administrator privileges
    - Removes file type registration for all users (HKEY_CLASSES_ROOT)

.EXAMPLE
    .\Unregister-GRDPFileType.ps1
#>

# Requires Administrator privileges
#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$fileExtension = ".grdp"
$progId = "GRIPS.DirectPrint.Archive"

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "GRIPS Direct Print File Type Unregistration" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Unregistering file type with the following settings:" -ForegroundColor White
Write-Host "  File Extension: $fileExtension" -ForegroundColor Gray
Write-Host "  ProgID: $progId" -ForegroundColor Gray
Write-Host ""

try {
    # Ensure HKCR: drive is available
    if (-not (Test-Path "HKCR:")) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    }

    $removed = $false

    # Remove the file extension registration
    $extKey = "HKCR:\$fileExtension"
    if (Test-Path $extKey) {
        Write-Host "1. Removing file extension registration..." -ForegroundColor Green
        Remove-Item -Path $extKey -Recurse -Force
        Write-Host "   ✓ Extension unregistered" -ForegroundColor Green
        $removed = $true
    } else {
        Write-Host "1. File extension not registered (skipping)" -ForegroundColor Yellow
    }
    Write-Host ""

    # Remove the ProgID registration
    $progIdKey = "HKCR:\$progId"
    if (Test-Path $progIdKey) {
        Write-Host "2. Removing ProgID registration..." -ForegroundColor Green
        Remove-Item -Path $progIdKey -Recurse -Force
        Write-Host "   ✓ ProgID unregistered" -ForegroundColor Green
        $removed = $true
    } else {
        Write-Host "2. ProgID not registered (skipping)" -ForegroundColor Yellow
    }
    Write-Host ""

    if ($removed) {
        # Notify Windows Explorer that file associations have changed
        Write-Host "3. Notifying Windows Explorer of changes..." -ForegroundColor Green
        
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
        
        Write-Host "   ✓ Explorer notified" -ForegroundColor Green
        Write-Host ""
    }

    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Unregistration completed successfully!" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($removed) {
        Write-Host "The .grdp file type has been unregistered from the system." -ForegroundColor White
    } else {
        Write-Host "No .grdp file type registration was found on the system." -ForegroundColor White
    }
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "ERROR: Unregistration failed!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    exit 1
}
