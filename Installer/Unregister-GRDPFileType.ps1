<#
.SYNOPSIS
    Unregisters the .grdp file type from Windows.

.DESCRIPTION
    This script removes the .grdp file extension registration from the Windows registry.
    By default, it only removes system-wide registrations (HKEY_CLASSES_ROOT) and warns
    about user-specific settings without removing them.

.PARAMETER IncludeLocal
    If specified, also attempts to remove user-specific file associations.
    Note: Windows protects UserChoice keys and they cannot be removed programmatically.

.NOTES
    - Must be run with Administrator privileges
    - Removes file type registration for all users (HKEY_CLASSES_ROOT)
    - User-specific associations (HKCU) are preserved by default

.EXAMPLE
    .\Unregister-GRDPFileType.ps1
    Removes system-wide registration only, warns about user settings.

.EXAMPLE
    .\Unregister-GRDPFileType.ps1 -IncludeLocal
    Removes system-wide registration and attempts to clean user-specific settings.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$IncludeLocal
)

# Requires Administrator privileges
#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$fileExtension = ".grdp"
$progId = "GRIPS.DirectPrint.Archive"
$mimeType = "application/x-grdp-archive"

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "GRIPS Direct Print File Type Unregistration" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Unregistering file type with the following settings:" -ForegroundColor White
Write-Host "  File Extension: $fileExtension" -ForegroundColor Gray
Write-Host "  ProgID: $progId" -ForegroundColor Gray
Write-Host "  MIME Type: $mimeType" -ForegroundColor Gray
if ($IncludeLocal) {
    Write-Host "  Mode: System-wide + Local user settings" -ForegroundColor Gray
} else {
    Write-Host "  Mode: System-wide only (use -IncludeLocal to remove user settings)" -ForegroundColor Gray
}
Write-Host ""

try {
    # Ensure HKCR: drive is available
    if (-not (Test-Path "HKCR:")) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    }

    $removed = $false

    # Remove the file extension registration from HKCR
    Write-Host "1. Removing system-wide file extension registration..." -ForegroundColor Green
    $extKeyPath = $fileExtension.TrimStart('.')
    $extRegKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($fileExtension, $false)
    
    if ($null -ne $extRegKey) {
        $extRegKey.Close()
        [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree($fileExtension)
        Write-Host "   [OK] Extension unregistered from HKCR" -ForegroundColor Green
        $removed = $true
    } else {
        Write-Host "   File extension not registered in HKCR (skipping)" -ForegroundColor Yellow
    }

    # Remove any auto_file associations
    $autoFileKeyPath = "${fileExtension}_auto_file"
    $autoFileRegKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($autoFileKeyPath, $false)
    if ($null -ne $autoFileRegKey) {
        $autoFileRegKey.Close()
        Write-Host "   Removing auto_file association..." -ForegroundColor Yellow
        [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree($autoFileKeyPath)
        Write-Host "   [OK] Auto_file removed" -ForegroundColor Green
        $removed = $true
    }
    Write-Host ""
    
    # Check for UserChoice keys (Windows protects these - cannot be removed programmatically)
    Write-Host "2. Checking for local user file associations..." -ForegroundColor Green
    $hasUserChoice = $false
    $hasLocalExtension = $false
    $hasLocalProgId = $false
    
    # Check HKEY_CURRENT_USER for local file extension registration
    $hkcuExtKey = "HKCU:\Software\Classes\$fileExtension"
    if (Test-Path $hkcuExtKey) {
        $hasLocalExtension = $true
        $localProgId = (Get-ItemProperty -Path $hkcuExtKey -Name "(Default)" -ErrorAction SilentlyContinue).'(Default)'
        Write-Host "   Found local extension registration in HKCU: $localProgId" -ForegroundColor Yellow
    }
    
    # Check HKEY_CURRENT_USER for local ProgID registration
    $hkcuProgIdKey = "HKCU:\Software\Classes\$progId"
    if (Test-Path $hkcuProgIdKey) {
        $hasLocalProgId = $true
        Write-Host "   Found local ProgID registration in HKCU: $progId" -ForegroundColor Yellow
    }
    
    # HKEY_CURRENT_USER UserChoice
    $userChoiceKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$fileExtension\UserChoice"
    if (Test-Path $userChoiceKey) {
        $currentChoice = (Get-ItemProperty -Path $userChoiceKey -Name "ProgId" -ErrorAction SilentlyContinue).ProgId
        Write-Host "   Found UserChoice for current user: $currentChoice" -ForegroundColor Yellow
        $hasUserChoice = $true
    }
    
    # Check OpenWithProgids
    $openWithKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$fileExtension\OpenWithProgids"
    if (Test-Path $openWithKey) {
        $progIds = Get-Item $openWithKey | Select-Object -ExpandProperty Property
        if ($progIds -contains $progId) {
            Write-Host "   Found OpenWithProgids reference for current user: $progId" -ForegroundColor Yellow
        }
    }
    
    # HKEY_USERS for all loaded profiles
    $hkuPath = "Registry::HKEY_USERS"
    if (Test-Path $hkuPath) {
        Get-ChildItem -Path $hkuPath -ErrorAction SilentlyContinue | ForEach-Object {
            $sid = $_.PSChildName
            
            # Check for local extension registration
            $userExtPath = Join-Path $_.PSPath "Software\Classes\$fileExtension"
            if (Test-Path $userExtPath) {
                $userProgId = (Get-ItemProperty -Path $userExtPath -Name "(Default)" -ErrorAction SilentlyContinue).'(Default)'
                Write-Host "   Found local extension for SID ${sid}: $userProgId" -ForegroundColor Yellow
                $hasLocalExtension = $true
            }
            
            # Check for local ProgID registration
            $userProgIdPath = Join-Path $_.PSPath "Software\Classes\$progId"
            if (Test-Path $userProgIdPath) {
                Write-Host "   Found local ProgID for SID ${sid}: $progId" -ForegroundColor Yellow
                $hasLocalProgId = $true
            }
            
            # Check for UserChoice
            $userChoicePath = Join-Path $_.PSPath "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$fileExtension\UserChoice"
            if (Test-Path $userChoicePath) {
                $profileChoice = (Get-ItemProperty -Path $userChoicePath -Name "ProgId" -ErrorAction SilentlyContinue).ProgId
                Write-Host "   Found UserChoice for SID ${sid}: $profileChoice" -ForegroundColor Yellow
                $hasUserChoice = $true
            }
        }
    }
    
    if ($hasLocalExtension -or $hasUserChoice -or $hasLocalProgId) {
        Write-Host ""
        if (-not $IncludeLocal) {
            Write-Host "   WARNING: Local user file associations found but NOT removed" -ForegroundColor Yellow
            Write-Host "   The system-wide registration will be removed, but user-specific" -ForegroundColor Yellow
            Write-Host "   settings are preserved to avoid breaking user file associations." -ForegroundColor Yellow
            Write-Host "" 
            Write-Host "   To remove ALL user settings, run: .\Unregister-GRDPFileType.ps1 -IncludeLocal" -ForegroundColor Cyan
            Write-Host "   Or manually: Right-click a .grdp file > Open with > Choose another app" -ForegroundColor Cyan
        } else {
            Write-Host "   Attempting to clean up local user settings..." -ForegroundColor Yellow
            
            # Remove HKCU extension registration
            if (Test-Path $hkcuExtKey) {
                Remove-Item -Path $hkcuExtKey -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "   [OK] Removed HKCU extension registration" -ForegroundColor Green
            }
            
            # Remove HKCU ProgID registration
            if (Test-Path $hkcuProgIdKey) {
                Remove-Item -Path $hkcuProgIdKey -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "   [OK] Removed HKCU ProgID registration" -ForegroundColor Green
            }
            
            # Note about UserChoice
            if ($hasUserChoice) {
                Write-Host "   NOTE: UserChoice keys cannot be removed (Windows protected)" -ForegroundColor Yellow
                Write-Host "   Users must manually change file association if needed" -ForegroundColor Yellow
            }
        }
        Write-Host "" 
    } else {
        Write-Host "   No local user associations found" -ForegroundColor Green
    }
    Write-Host ""

    # Remove the ProgID registration from HKCR (only if not also registered locally)
    Write-Host "3. Removing system-wide ProgID registration..." -ForegroundColor Green
    
    if ($hasLocalProgId -and -not $IncludeLocal) {
        Write-Host "   ProgID is registered locally, skipping HKCR removal to preserve user associations" -ForegroundColor Yellow
    } else {
        $progIdRegKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($progId, $false)
        
        if ($null -ne $progIdRegKey) {
            $progIdRegKey.Close()
            [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree($progId)
            Write-Host "   [OK] ProgID unregistered from HKCR" -ForegroundColor Green
            $removed = $true
        } else {
            Write-Host "   ProgID not registered in HKCR (skipping)" -ForegroundColor Yellow
        }
    }
    Write-Host ""

    # Remove the MIME type registration
    Write-Host "4. Removing MIME type registration..." -ForegroundColor Green
    $mimeDbKeyPath = "MIME\Database\Content Type\$mimeType"
    $mimeRemoved = $false
    
    try {
        # Try to get the key from the registry directly (read-only first to check existence)
        $regKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($mimeDbKeyPath, $false)
        if ($null -ne $regKey) {
            $regKey.Close()
            # Key exists, delete it using the .NET API
            [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree($mimeDbKeyPath)
            Write-Host "   [OK] MIME type unregistered" -ForegroundColor Green
            $removed = $true
            $mimeRemoved = $true
        }
    } catch {
        # Key doesn't exist or couldn't be deleted
    }
    
    # Also check for incorrectly created nested keys (if PowerShell treated / as path separator)
    # This would create: MIME\Database\Content Type\application\x-grdp-archive instead of
    # MIME\Database\Content Type\application/x-grdp-archive
    try {
        $incorrectPath = "MIME\Database\Content Type\application"
        $appKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($incorrectPath, $false)
        if ($null -ne $appKey) {
            $subKeys = $appKey.GetSubKeyNames()
            $appKey.Close()
            if ($subKeys -contains "x-grdp-archive") {
                Write-Host "   Removing incorrectly nested MIME type keys..." -ForegroundColor Yellow
                [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree("$incorrectPath\x-grdp-archive")
                $removed = $true
                $mimeRemoved = $true
            }
        }
    } catch {
        # Nested keys don't exist
    }
    
    if ($mimeRemoved) {
        # Don't show OK again since we showed it above
    } else {
        Write-Host "   MIME type not registered (skipping)" -ForegroundColor Yellow
    }
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
