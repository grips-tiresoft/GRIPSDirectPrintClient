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
$mimeType = "application/x-grdp-archive"

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
Write-Host "  MIME Type: $mimeType" -ForegroundColor Gray
Write-Host ""

try {
    # Ensure HKCR: drive is available
    if (-not (Test-Path "HKCR:")) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    }

    # Clean up any existing auto_file associations that might interfere
    Write-Host "1. Cleaning up existing associations..." -ForegroundColor Green
    $autoFileKeyPath = "${fileExtension}_auto_file"
    $autoFileRegKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($autoFileKeyPath, $false)
    if ($null -ne $autoFileRegKey) {
        $autoFileRegKey.Close()
        Write-Host "   Removing ${fileExtension}_auto_file..." -ForegroundColor Yellow
        [Microsoft.Win32.Registry]::ClassesRoot.DeleteSubKeyTree($autoFileKeyPath)
    }
    
    # Check for UserChoice keys (these override system defaults and CANNOT be removed programmatically)
    $hasUserChoice = $false
    $userChoiceKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$fileExtension\UserChoice"
    if (Test-Path $userChoiceKey) {
        $hasUserChoice = $true
        $currentChoice = (Get-ItemProperty -Path $userChoiceKey -Name "ProgId" -ErrorAction SilentlyContinue).ProgId
        Write-Host "   WARNING: User has manually selected an app to open $fileExtension files: $currentChoice" -ForegroundColor Yellow
    }
    
    # Also check HKEY_USERS for all loaded user profiles
    $hkuPath = "Registry::HKEY_USERS"
    if (Test-Path $hkuPath) {
        Get-ChildItem -Path $hkuPath -ErrorAction SilentlyContinue | ForEach-Object {
            $userChoicePath = Join-Path $_.PSPath "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$fileExtension\UserChoice"
            if (Test-Path $userChoicePath) {
                $hasUserChoice = $true
                $profileChoice = (Get-ItemProperty -Path $userChoicePath -Name "ProgId" -ErrorAction SilentlyContinue).ProgId
                Write-Host "   WARNING: User profile $($_.PSChildName) has manual app selection: $profileChoice" -ForegroundColor Yellow
            }
        }
    }
    
    if ($hasUserChoice) {
        Write-Host "" 
        Write-Host "   IMPORTANT: Windows protects user app choices. To use GRIPS Direct Print:" -ForegroundColor Yellow
        Write-Host "   1. Right-click a .grdp file in Explorer" -ForegroundColor Yellow
        Write-Host "   2. Select 'Open with' > 'Choose another app'" -ForegroundColor Yellow
        Write-Host "   3. Select 'GRIPS Direct Print' from the list" -ForegroundColor Yellow
        Write-Host "   4. Check 'Always use this app' and click OK" -ForegroundColor Yellow
        Write-Host "" 
    }
    Write-Host "   [OK] Cleanup completed" -ForegroundColor Green
    Write-Host ""

    # Register the file extension
    Write-Host "2. Registering file extension $fileExtension..." -ForegroundColor Green
    $extKeyPath = "$fileExtension"
    $extRegKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($extKeyPath, $false)
    
    if ($null -ne $extRegKey) {
        $extRegKey.Close()
        Write-Host "   Extension already registered, updating..." -ForegroundColor Yellow
    }
    
    $extKey = "HKCR:\$fileExtension"
    New-Item -Path $extKey -Force | Out-Null
    New-ItemProperty -Path $extKey -Name "(Default)" -Value $progId -Force | Out-Null
    
    # Set the PerceivedType to help Windows understand this is a document type
    New-ItemProperty -Path $extKey -Name "PerceivedType" -Value "document" -Force | Out-Null
    
    # Set the MIME Content Type
    New-ItemProperty -Path $extKey -Name "Content Type" -Value $mimeType -Force | Out-Null
    
    Write-Host "   [OK] Extension registered" -ForegroundColor Green
    Write-Host ""

    # Register the MIME type in the MIME Database
    Write-Host "3. Registering MIME type $mimeType..." -ForegroundColor Green
    
    # Use .NET Registry API to create the key with forward slash in the name
    # PowerShell's New-Item treats forward slashes as path separators, which is incorrect for MIME types
    $contentTypeKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("MIME\Database\Content Type", $true)
    $mimeKey = $contentTypeKey.CreateSubKey($mimeType)
    $mimeKey.SetValue("Extension", $fileExtension, [Microsoft.Win32.RegistryValueKind]::String)
    $mimeKey.Close()
    $contentTypeKey.Close()
    
    Write-Host "   [OK] MIME type registered" -ForegroundColor Green
    Write-Host ""

    # Register the ProgID
    Write-Host "4. Registering ProgID $progId..." -ForegroundColor Green
    $progIdKey = "HKCR:\$progId"
    
    New-Item -Path $progIdKey -Force | Out-Null
    New-ItemProperty -Path $progIdKey -Name "(Default)" -Value $friendlyTypeName -Force | Out-Null
    New-ItemProperty -Path $progIdKey -Name "FriendlyTypeName" -Value $friendlyTypeName -Force | Out-Null
    Write-Host "   [OK] ProgID registered" -ForegroundColor Green
    Write-Host ""

    # Set the default icon
    Write-Host "5. Setting default icon..." -ForegroundColor Green
    $iconKey = "$progIdKey\DefaultIcon"
    New-Item -Path $iconKey -Force | Out-Null
    New-ItemProperty -Path $iconKey -Name "(Default)" -Value $iconPath -Force | Out-Null
    Write-Host "   [OK] Icon set" -ForegroundColor Green
    Write-Host ""

    # Register the shell command to open the file
    Write-Host "6. Registering shell open command..." -ForegroundColor Green
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
    Write-Host "7. Notifying Windows Explorer of changes..." -ForegroundColor Green
    
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
sleep 2
# SIG # Begin signature block
# MII7sgYJKoZIhvcNAQcCoII7ozCCO58CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB8Yp7iAZJQ2ZK3
# /TShQvo+oKKuFvemLbHHhoLi3VZ40KCCI9YwggXMMIIDtKADAgECAhBUmNLR1FsZ
# lUgTecgRwIeZMA0GCSqGSIb3DQEBDAUAMHcxCzAJBgNVBAYTAlVTMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xSDBGBgNVBAMTP01pY3Jvc29mdCBJZGVu
# dGl0eSBWZXJpZmljYXRpb24gUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAy
# MDAeFw0yMDA0MTYxODM2MTZaFw00NTA0MTYxODQ0NDBaMHcxCzAJBgNVBAYTAlVT
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xSDBGBgNVBAMTP01pY3Jv
# c29mdCBJZGVudGl0eSBWZXJpZmljYXRpb24gUm9vdCBDZXJ0aWZpY2F0ZSBBdXRo
# b3JpdHkgMjAyMDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALORKgeD
# Bmf9np3gx8C3pOZCBH8Ppttf+9Va10Wg+3cL8IDzpm1aTXlT2KCGhFdFIMeiVPvH
# or+Kx24186IVxC9O40qFlkkN/76Z2BT2vCcH7kKbK/ULkgbk/WkTZaiRcvKYhOuD
# PQ7k13ESSCHLDe32R0m3m/nJxxe2hE//uKya13NnSYXjhr03QNAlhtTetcJtYmrV
# qXi8LW9J+eVsFBT9FMfTZRY33stuvF4pjf1imxUs1gXmuYkyM6Nix9fWUmcIxC70
# ViueC4fM7Ke0pqrrBc0ZV6U6CwQnHJFnni1iLS8evtrAIMsEGcoz+4m+mOJyoHI1
# vnnhnINv5G0Xb5DzPQCGdTiO0OBJmrvb0/gwytVXiGhNctO/bX9x2P29Da6SZEi3
# W295JrXNm5UhhNHvDzI9e1eM80UHTHzgXhgONXaLbZ7LNnSrBfjgc10yVpRnlyUK
# xjU9lJfnwUSLgP3B+PR0GeUw9gb7IVc+BhyLaxWGJ0l7gpPKWeh1R+g/OPTHU3mg
# trTiXFHvvV84wRPmeAyVWi7FQFkozA8kwOy6CXcjmTimthzax7ogttc32H83rwjj
# O3HbbnMbfZlysOSGM1l0tRYAe1BtxoYT2v3EOYI9JACaYNq6lMAFUSw0rFCZE4e7
# swWAsk0wAly4JoNdtGNz764jlU9gKL431VulAgMBAAGjVDBSMA4GA1UdDwEB/wQE
# AwIBhjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTIftJqhSobyhmYBAcnz1AQ
# T2ioojAQBgkrBgEEAYI3FQEEAwIBADANBgkqhkiG9w0BAQwFAAOCAgEAr2rd5hnn
# LZRDGU7L6VCVZKUDkQKL4jaAOxWiUsIWGbZqWl10QzD0m/9gdAmxIR6QFm3FJI9c
# Zohj9E/MffISTEAQiwGf2qnIrvKVG8+dBetJPnSgaFvlVixlHIJ+U9pW2UYXeZJF
# xBA2CFIpF8svpvJ+1Gkkih6PsHMNzBxKq7Kq7aeRYwFkIqgyuH4yKLNncy2RtNwx
# AQv3Rwqm8ddK7VZgxCwIo3tAsLx0J1KH1r6I3TeKiW5niB31yV2g/rarOoDXGpc8
# FzYiQR6sTdWD5jw4vU8w6VSp07YEwzJ2YbuwGMUrGLPAgNW3lbBeUU0i/OxYqujY
# lLSlLu2S3ucYfCFX3VVj979tzR/SpncocMfiWzpbCNJbTsgAlrPhgzavhgplXHT2
# 6ux6anSg8Evu75SjrFDyh+3XOjCDyft9V77l4/hByuVkrrOj7FjshZrM77nq81YY
# uVxzmq/FdxeDWds3GhhyVKVB0rYjdaNDmuV3fJZ5t0GNv+zcgKCf0Xd1WF81E+Al
# GmcLfc4l+gcK5GEh2NQc5QfGNpn0ltDGFf5Ozdeui53bFv0ExpK91IjmqaOqu/dk
# ODtfzAzQNb50GQOmxapMomE2gj4d8yu8l13bS3g7LfU772Aj6PXsCyM2la+YZr9T
# 03u4aUoqlmZpxJTG9F9urJh4iIAGXKKy7aIwggdaMIIFQqADAgECAhMzAAAABkoa
# +s8FYWp0AAAAAAAGMA0GCSqGSIb3DQEBDAUAMGMxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xNDAyBgNVBAMTK01pY3Jvc29mdCBJ
# RCBWZXJpZmllZCBDb2RlIFNpZ25pbmcgUENBIDIwMjEwHhcNMjEwNDEzMTczMTU0
# WhcNMjYwNDEzMTczMTU0WjBaMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSswKQYDVQQDEyJNaWNyb3NvZnQgSUQgVmVyaWZpZWQg
# Q1MgRU9DIENBIDAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAx+PI
# P/Qh3cYZwLvFy6uuJ4fTp3ln7Gqs7s8lTVyfgOJWP1aABwk2/oxdVjfSHUq4MTPX
# ilL57qi/fH7YndEK4Knd3u5cedFwr2aHSTp6vl/PL1dAL9sfoDvNpdG0N/R84AhY
# NpBQThpO4/BqxmCgl3iIRfhh2oFVOuiTiDVWvXBg76bcjnHnEEtXzvAWwJu0bBU7
# oRRqQed4VXJtICVt+ZoKUSjqY5wUlhAdwHh+31BnpBPCzFtKViLp6zEtRyOxRega
# gFU+yLgXvvmd07IDN0S2TLYuiZjTw+kcYOtoNgKr7k0C6E9Wf3H4jHavk2MxqFpt
# gfL0gL+zbSb+VBNKiVT0mqzXJIJmWmqw0K+D3MKfmCer3e3CbrP+F5RtCb0XaE0u
# RcJPZJjWwciDBxBIbkNF4GL12hl5vydgFMmzQcNuodKyX//3lLJ1q22roHVS1cgt
# sLgpjWYZlBlhCTcXJeZ3xuaJvXZB9rcLCX15OgXL21tUUwJCLE27V5AGZxkO3i54
# mgSCswtOmWU4AKd/B/e3KtXv6XBURKuAteez1EpgloaZwQej9l5dN9Uh8W19BZg9
# IlLl+xHRX4vDiMWAUf/7ANe4MoS98F45r76IGJ0hC02EMuMZxAErwZj0ln0aL53E
# zlMa5JCiRObb0UoLHfGSdNJsMg0uj3DAQDdVWTECAwEAAaOCAg4wggIKMA4GA1Ud
# DwEB/wQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUdpw2dBPRkH1h
# X7MC64D0mUulPoUwVAYDVR0gBE0wSzBJBgRVHSAAMEEwPwYIKwYBBQUHAgEWM2h0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0
# bTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB8GA1UdIwQYMBaAFNlBKbAPD2Ns72nX9c0pnqRIajDmMHAGA1UdHwRpMGcwZaBj
# oGGGX2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29m
# dCUyMElEJTIwVmVyaWZpZWQlMjBDb2RlJTIwU2lnbmluZyUyMFBDQSUyMDIwMjEu
# Y3JsMIGuBggrBgEFBQcBAQSBoTCBnjBtBggrBgEFBQcwAoZhaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBJRCUyMFZlcmlm
# aWVkJTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDIxLmNydDAtBggrBgEFBQcw
# AYYhaHR0cDovL29uZW9jc3AubWljcm9zb2Z0LmNvbS9vY3NwMA0GCSqGSIb3DQEB
# DAUAA4ICAQBqLwmf2LB1QjUga0G7zFkbGd8NBQLHP0KOFBWNJFZiTtKfpO0bZ2Wf
# s6v5vqIKjE32Q6M89G4ZkVcvWuEAA+dvjLThSy89Y0//m/WTSKwYtiR1Ewn7x1kw
# /Fg93wQps2C1WUj+00/6uNrF+d4MVJxV1HoBID+95ZIW0KkqZopnOA4w5vP4T5cB
# prZQAlP/vMGyB0H9+pHNo0jT9Q8gfKJNzHS9i1DgBmmufGdW9TByuno8GAizFMhL
# lIs08b5lilIkE5z3FMAUAr+XgII1FNZnb43OI6Qd2zOijbjYfursXUCNHC+RSwJG
# m5ULzPymYggnJ+khJOq7oSlqPGpbr70hGBePw/J7/mmSqp7hTgt0mPikS1i4ap8x
# +P3yemYShnFrgV1752TI+As69LfgLthkITvf7bFHB8vmIhadZCOS0vTCx3B+/OVc
# EMLNO2bJ0O9ikc1JqR0Fvqx7nAwMRSh3FVqosgzBbWnVkQJq7oWFwMVfFIYn6LPR
# ZMt48u6iMUCFBSPddsPA/6k85mEv+08U5WCQ7ydj1KVV2THre/8mLHiem9wf/Czo
# hqRntxM2E/x+NHy6TBMnSPQRqhhNfuOgUDAWEYmlM/ZHGaPIb7xOvfVyLQ/7l6Yf
# ogT3eptwp4GOGRjH5z+gG9kpBIx8QrRl6OilnlxRExokmMflL7l12TCCB38wggVn
# oAMCAQICEzMABrMAX86rBzNERTYAAAAGswAwDQYJKoZIhvcNAQEMBQAwWjELMAkG
# A1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjErMCkGA1UE
# AxMiTWljcm9zb2Z0IElEIFZlcmlmaWVkIENTIEVPQyBDQSAwMTAeFw0yNjAxMjgx
# NTA4MzBaFw0yNjAxMzExNTA4MzBaMIH8MRMwEQYDVQQREwo0NDMxNi0wMDAxMQsw
# CQYDVQQGEwJVUzENMAsGA1UECBMET2hpbzEOMAwGA1UEBxMFQWtyb24xGzAZBgNV
# BAkTEjIwMCBJbm5vdmF0aW9uIFdheTFNMEsGA1UECh5EAFQAaABlACAARwBvAG8A
# ZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBtAHAA
# YQBuAHkxTTBLBgNVBAMeRABUAGgAZQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQAaQBy
# AGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5MIIBojANBgkqhkiG
# 9w0BAQEFAAOCAY8AMIIBigKCAYEAjjWpruG30itDrQsIv68L+CQaXESK2RGV90Pn
# 322F67CjmloCdXE5n+PIl0bcERNYbbjgbt9MQqWw6M6TleGz2FYbvnRxyzuSv+jZ
# xtFcM/9v5MVt7hxDNKpxb+lec4KF2DR4Sm9vHkTRQOxKsznBVGoBG0no5h6l02mF
# KI2B6TPAt203fqepxDGQasZPxto+vvDAN1tjYApgAEont4KS96kIdB6wzGJ+wUaN
# 1io9QsaKu5f0K+mTx4e7kWsLGRsWi0wqAL8Hca9iFPQeeSXzj1WD6jrSr22kxWbE
# 6fTMCQBFu51ftI6xNRP0g9c4jZJBlRQX5iLZG7eeKldSrKX4I/en2IISzW5dd09+
# QAwOdGcSkVLh2yQTOb/ZEhKSDpWrjX/JB1bR/rl8khsPyefpDk46JrGqFaZKmZsH
# fGNLP6axvvYwqMOrx66sKkoqlqNGuwRH89C+U2bt+IyJ+M29D0PPxIslZTkzo3kI
# qr+vcuQW5afCjLZkoo63oy2qOwFRAgMBAAGjggIZMIICFTAMBgNVHRMBAf8EAjAA
# MA4GA1UdDwEB/wQEAwIHgDA8BgNVHSUENTAzBgorBgEEAYI3YQEABggrBgEFBQcD
# AwYbKwYBBAGCN2GDi6zmXoHCg+Bng6e09AGZrMsTMB0GA1UdDgQWBBTSLNHyDJQv
# LRlvPjuMD6ranWTpUDAfBgNVHSMEGDAWgBR2nDZ0E9GQfWFfswLrgPSZS6U+hTBn
# BgNVHR8EYDBeMFygWqBYhlZodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBFT0MlMjBDQSUy
# MDAxLmNybDCBpQYIKwYBBQUHAQEEgZgwgZUwZAYIKwYBBQUHMAKGWGh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSUQlMjBW
# ZXJpZmllZCUyMENTJTIwRU9DJTIwQ0ElMjAwMS5jcnQwLQYIKwYBBQUHMAGGIWh0
# dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDBmBgNVHSAEXzBdMFEGDCsG
# AQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQBMA0GCSqGSIb3
# DQEBDAUAA4ICAQANZ6gprHiJS41WiPRldB7FNMswXfal+hXoP+l8CBbSf8BKSBeL
# WGqGMfrq0BfjLtE76J9ujoRvBaP1z1kmhB1hds8Q3yF4letugz0IFyn6zvDET5Cq
# VATnsZHzdTqwRN5oBmtqG1tjapDAuG7qi6vQaBTOkpDry0hkL68wBPhKqE+rBaWV
# tm68IecylXAnWMXlEewWO5HBzEBw+Akm9CuW/4TjNNQ4mx7upS8T84wqSYpf4jBx
# qhUQqCIpmfZH1laXhIevZJSRgStJ7UaVttiz7KPke9wIZ8DbcKhsRDX5/M2w4BCP
# Nw0W7SpDgwSzypgp7UI1mYsW4KHNAy+STUmVKbCdjol/rfma3HVaAOCqELv30aHE
# ClC4D9nk5+gkZRxLPXo4VAyNJoY4AjKKvN03r8cdF6y0rW/wIfgp8xIc3A/XXMLQ
# SwgYAiAI+XIdNyCWhGAEHIOAaw8EnvwaqSmydsaoN5wqQSB2N6BV3RMucxCLiGVZ
# CpcWfi/dMJuVOt+wJj7nN3ZQ9O3snMWJ9ynMY5e/t7qjLjoH5fuE9v5ky+I0BI2Y
# o3mhLaHZcL4zNrUkb3vRNNoO3ViAa6oLP8EbtM5ro+9FajgBLwX9MZjtTrPKJMPR
# 70xX07yEGfAgYCjdC9hT00HsbofXmHo7Q/fDl+6wqBnQpwzhkqIR4wwrXjCCB38w
# ggVnoAMCAQICEzMABrMAX86rBzNERTYAAAAGswAwDQYJKoZIhvcNAQEMBQAwWjEL
# MAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjErMCkG
# A1UEAxMiTWljcm9zb2Z0IElEIFZlcmlmaWVkIENTIEVPQyBDQSAwMTAeFw0yNjAx
# MjgxNTA4MzBaFw0yNjAxMzExNTA4MzBaMIH8MRMwEQYDVQQREwo0NDMxNi0wMDAx
# MQswCQYDVQQGEwJVUzENMAsGA1UECBMET2hpbzEOMAwGA1UEBxMFQWtyb24xGzAZ
# BgNVBAkTEjIwMCBJbm5vdmF0aW9uIFdheTFNMEsGA1UECh5EAFQAaABlACAARwBv
# AG8AZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBt
# AHAAYQBuAHkxTTBLBgNVBAMeRABUAGgAZQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQA
# aQByAGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5MIIBojANBgkq
# hkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAjjWpruG30itDrQsIv68L+CQaXESK2RGV
# 90Pn322F67CjmloCdXE5n+PIl0bcERNYbbjgbt9MQqWw6M6TleGz2FYbvnRxyzuS
# v+jZxtFcM/9v5MVt7hxDNKpxb+lec4KF2DR4Sm9vHkTRQOxKsznBVGoBG0no5h6l
# 02mFKI2B6TPAt203fqepxDGQasZPxto+vvDAN1tjYApgAEont4KS96kIdB6wzGJ+
# wUaN1io9QsaKu5f0K+mTx4e7kWsLGRsWi0wqAL8Hca9iFPQeeSXzj1WD6jrSr22k
# xWbE6fTMCQBFu51ftI6xNRP0g9c4jZJBlRQX5iLZG7eeKldSrKX4I/en2IISzW5d
# d09+QAwOdGcSkVLh2yQTOb/ZEhKSDpWrjX/JB1bR/rl8khsPyefpDk46JrGqFaZK
# mZsHfGNLP6axvvYwqMOrx66sKkoqlqNGuwRH89C+U2bt+IyJ+M29D0PPxIslZTkz
# o3kIqr+vcuQW5afCjLZkoo63oy2qOwFRAgMBAAGjggIZMIICFTAMBgNVHRMBAf8E
# AjAAMA4GA1UdDwEB/wQEAwIHgDA8BgNVHSUENTAzBgorBgEEAYI3YQEABggrBgEF
# BQcDAwYbKwYBBAGCN2GDi6zmXoHCg+Bng6e09AGZrMsTMB0GA1UdDgQWBBTSLNHy
# DJQvLRlvPjuMD6ranWTpUDAfBgNVHSMEGDAWgBR2nDZ0E9GQfWFfswLrgPSZS6U+
# hTBnBgNVHR8EYDBeMFygWqBYhlZodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBFT0MlMjBD
# QSUyMDAxLmNybDCBpQYIKwYBBQUHAQEEgZgwgZUwZAYIKwYBBQUHMAKGWGh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSUQl
# MjBWZXJpZmllZCUyMENTJTIwRU9DJTIwQ0ElMjAwMS5jcnQwLQYIKwYBBQUHMAGG
# IWh0dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDBmBgNVHSAEXzBdMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQBMA0GCSqG
# SIb3DQEBDAUAA4ICAQANZ6gprHiJS41WiPRldB7FNMswXfal+hXoP+l8CBbSf8BK
# SBeLWGqGMfrq0BfjLtE76J9ujoRvBaP1z1kmhB1hds8Q3yF4letugz0IFyn6zvDE
# T5CqVATnsZHzdTqwRN5oBmtqG1tjapDAuG7qi6vQaBTOkpDry0hkL68wBPhKqE+r
# BaWVtm68IecylXAnWMXlEewWO5HBzEBw+Akm9CuW/4TjNNQ4mx7upS8T84wqSYpf
# 4jBxqhUQqCIpmfZH1laXhIevZJSRgStJ7UaVttiz7KPke9wIZ8DbcKhsRDX5/M2w
# 4BCPNw0W7SpDgwSzypgp7UI1mYsW4KHNAy+STUmVKbCdjol/rfma3HVaAOCqELv3
# 0aHEClC4D9nk5+gkZRxLPXo4VAyNJoY4AjKKvN03r8cdF6y0rW/wIfgp8xIc3A/X
# XMLQSwgYAiAI+XIdNyCWhGAEHIOAaw8EnvwaqSmydsaoN5wqQSB2N6BV3RMucxCL
# iGVZCpcWfi/dMJuVOt+wJj7nN3ZQ9O3snMWJ9ynMY5e/t7qjLjoH5fuE9v5ky+I0
# BI2Yo3mhLaHZcL4zNrUkb3vRNNoO3ViAa6oLP8EbtM5ro+9FajgBLwX9MZjtTrPK
# JMPR70xX07yEGfAgYCjdC9hT00HsbofXmHo7Q/fDl+6wqBnQpwzhkqIR4wwrXjCC
# B54wggWGoAMCAQICEzMAAAAHh6M0o3uljhwAAAAAAAcwDQYJKoZIhvcNAQEMBQAw
# dzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjFI
# MEYGA1UEAxM/TWljcm9zb2Z0IElkZW50aXR5IFZlcmlmaWNhdGlvbiBSb290IENl
# cnRpZmljYXRlIEF1dGhvcml0eSAyMDIwMB4XDTIxMDQwMTIwMDUyMFoXDTM2MDQw
# MTIwMTUyMFowYzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjE0MDIGA1UEAxMrTWljcm9zb2Z0IElEIFZlcmlmaWVkIENvZGUgU2ln
# bmluZyBQQ0EgMjAyMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALLw
# wK8ZiCji3VR6TElsaQhVCbRS/3pK+MHrJSj3Zxd3KU3rlfL3qrZilYKJNqztA9OQ
# acr1AwoNcHbKBLbsQAhBnIB34zxf52bDpIO3NJlfIaTE/xrweLoQ71lzCHkD7A4A
# s1Bs076Iu+mA6cQzsYYH/Cbl1icwQ6C65rU4V9NQhNUwgrx9rGQ//h890Q8JdjLL
# w0nV+ayQ2Fbkd242o9kH82RZsH3HEyqjAB5a8+Ae2nPIPc8sZU6ZE7iRrRZywRmr
# KDp5+TcmJX9MRff241UaOBs4NmHOyke8oU1TYrkxh+YeHgfWo5tTgkoSMoayqoDp
# HOLJs+qG8Tvh8SnifW2Jj3+ii11TS8/FGngEaNAWrbyfNrC69oKpRQXY9bGH6jn9
# NEJv9weFxhTwyvx9OJLXmRGbAUXN1U9nf4lXezky6Uh/cgjkVd6CGUAf0K+Jw+GE
# /5VpIVbcNr9rNE50Sbmy/4RTCEGvOq3GhjITbCa4crCzTTHgYYjHs1NbOc6brH+e
# KpWLtr+bGecy9CrwQyx7S/BfYJ+ozst7+yZtG2wR461uckFu0t+gCwLdN0A6cFtS
# RtR8bvxVFyWwTtgMMFRuBa3vmUOTnfKLsLefRaQcVTgRnzeLzdpt32cdYKp+dhr2
# ogc+qM6K4CBI5/j4VFyC4QFeUP2YAidLtvpXRRo3AgMBAAGjggI1MIICMTAOBgNV
# HQ8BAf8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFNlBKbAPD2Ns
# 72nX9c0pnqRIajDmMFQGA1UdIARNMEswSQYEVR0gADBBMD8GCCsGAQUFBwIBFjNo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5o
# dG0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwDwYDVR0TAQH/BAUwAwEB/zAf
# BgNVHSMEGDAWgBTIftJqhSobyhmYBAcnz1AQT2ioojCBhAYDVR0fBH0wezB5oHeg
# dYZzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwSWRlbnRpdHklMjBWZXJpZmljYXRpb24lMjBSb290JTIwQ2VydGlmaWNhdGUl
# MjBBdXRob3JpdHklMjAyMDIwLmNybDCBwwYIKwYBBQUHAQEEgbYwgbMwgYEGCCsG
# AQUFBzAChnVodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01p
# Y3Jvc29mdCUyMElkZW50aXR5JTIwVmVyaWZpY2F0aW9uJTIwUm9vdCUyMENlcnRp
# ZmljYXRlJTIwQXV0aG9yaXR5JTIwMjAyMC5jcnQwLQYIKwYBBQUHMAGGIWh0dHA6
# Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDANBgkqhkiG9w0BAQwFAAOCAgEA
# fyUqnv7Uq+rdZgrbVyNMul5skONbhls5fccPlmIbzi+OwVdPQ4H55v7VOInnmezQ
# EeW4LqK0wja+fBznANbXLB0KrdMCbHQpbLvG6UA/Xv2pfpVIE1CRFfNF4XKO8XYE
# a3oW8oVH+KZHgIQRIwAbyFKQ9iyj4aOWeAzwk+f9E5StNp5T8FG7/VEURIVWArbA
# zPt9ThVN3w1fAZkF7+YU9kbq1bCR2YD+MtunSQ1Rft6XG7b4e0ejRA7mB2IoX5hN
# h3UEauY0byxNRG+fT2MCEhQl9g2i2fs6VOG19CNep7SquKaBjhWmirYyANb0RJSL
# WjinMLXNOAga10n8i9jqeprzSMU5ODmrMCJE12xS/NWShg/tuLjAsKP6SzYZ+1Ry
# 358ZTFcx0FS/mx2vSoU8s8HRvy+rnXqyUJ9HBqS0DErVLjQwK8VtsBdekBmdTbQV
# oCgPCqr+PDPB3xajYnzevs7eidBsM71PINK2BoE2UfMwxCCX3mccFgx6UsQeRSdV
# VVNSyALQe6PT12418xon2iDGE81OGCreLzDcMAZnrUAx4XQLUz6ZTl65yPUiOh3k
# 7Yww94lDf+8oG2oZmDh5O1Qe38E+M3vhKwmzIeoB1dVLlz4i3IpaDcR+iuGjH2Td
# aC1ZOmBXiCRKJLj4DT2uhJ04ji+tHD6n58vhavFIrmcxghcyMIIXLgIBATBxMFox
# CzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKzAp
# BgNVBAMTIk1pY3Jvc29mdCBJRCBWZXJpZmllZCBDUyBFT0MgQ0EgMDECEzMABrMA
# X86rBzNERTYAAAAGswAwDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3AgEMMQIw
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQg60im
# X/kmoW1S5hdgaxfZWEBNQbC8bFZEJlZkxrxtSrUwDQYJKoZIhvcNAQEBBQAEggGA
# cbfuIm22F1YVXf7mHfxX2oSiNxODljUlT6iM41/umOIMGeDMO6RCEY868f+ScfM9
# py3ixPJPVWvfexNfI6I0eae5ZEwK2T/oW+hRy4ReKsHxW9yT7gbCSvxLivml+FaQ
# /htNqVSR35z4uhvOC3wHlvfMH/Zz4JFfZjo9prsq3hTN7KXwCpYxaLC1YkQEUmmR
# QP/Jkb29RHjdvGlIMGO0j3vlQ/+cJJ26u0Wl/v6IPMU/x6qXu1TZSV/6rLgQkyyL
# O/7lXcr4qxlR1W3H01LRNy8o5SGddUxPu8XheGu6FsByKnklSmKNBCIv4xwQKLAp
# WnRCdVmz6WNzTZKdvN5nMq4GluxqGQlwBt0kNmZLbEZdQCkBlXjuxTuIqB6H/TtQ
# n7oLOblstjrJBtVOaXq+9U42n3fCrehy6P4oMMFmlVah51vtalpKXOApUeRDwtxf
# Mmbdz9axLSMiVIdGA7WbECJG/0M112C95VHHmmEwG8m1hCDgmUuRqIVn4Na253jv
# oYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIUizCCFIcC
# AQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSCAVUwggFR
# AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIOgLSoGdAZhTLRZ8y/CI
# mQ3zJagxAkrCdIZCnQtIOCZsAgZpePK6GHkYEzIwMjYwMTI5MTQxOTA3LjU2N1ow
# BIACAfSggemkgeYwgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRl
# ZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjdCMUEtMDVFMC1EOTQ3MTUwMwYD
# VQQDEyxNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1dGhvcml0
# eaCCDykwggeCMIIFaqADAgECAhMzAAAABeXPD/9mLsmHAAAAAAAFMA0GCSqGSIb3
# DQEBDAUAMHcxCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xSDBGBgNVBAMTP01pY3Jvc29mdCBJZGVudGl0eSBWZXJpZmljYXRpb24g
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAyMDAeFw0yMDExMTkyMDMyMzFa
# Fw0zNTExMTkyMDQyMzFaMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMgUlNBIFRp
# bWVzdGFtcGluZyBDQSAyMDIwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAnnznUmP94MWfBX1jtQYioxwe1+eXM9ETBb1lRkd3kcFdcG9/sqtDlwxKoVIc
# aqDb+omFio5DHC4RBcbyQHjXCwMk/l3TOYtgoBjxnG/eViS4sOx8y4gSq8Zg49RE
# Af5huXhIkQRKe3Qxs8Sgp02KHAznEa/Ssah8nWo5hJM1xznkRsFPu6rfDHeZeG1W
# a1wISvlkpOQooTULFm809Z0ZYlQ8Lp7i5F9YciFlyAKwn6yjN/kR4fkquUWfGmMo
# pNq/B8U/pdoZkZZQbxNlqJOiBGgCWpx69uKqKhTPVi3gVErnc/qi+dR8A2MiAz0k
# N0nh7SqINGbmw5OIRC0EsZ31WF3Uxp3GgZwetEKxLms73KG/Z+MkeuaVDQQheang
# OEMGJ4pQZH55ngI0Tdy1bi69INBV5Kn2HVJo9XxRYR/JPGAaM6xGl57Ei95HUw9N
# V/uC3yFjrhc087qLJQawSC3xzY/EXzsT4I7sDbxOmM2rl4uKK6eEpurRduOQ2hTk
# mG1hSuWYBunFGNv21Kt4N20AKmbeuSnGnsBCd2cjRKG79+TX+sTehawOoxfeOO/j
# R7wo3liwkGdzPJYHgnJ54UxbckF914AqHOiEV7xTnD1a69w/UTxwjEugpIPMIIE6
# 7SFZ2PMo27xjlLAHWW3l1CEAFjLNHd3EQ79PUr8FUXetXr0CAwEAAaOCAhswggIX
# MA4GA1UdDwEB/wQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUa2ko
# OjUvSGNAz3vYr0npPtk92yEwVAYDVR0gBE0wSzBJBgRVHSAAMEEwPwYIKwYBBQUH
# AgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
# b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMA
# dQBiAEMAQTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFMh+0mqFKhvKGZgE
# ByfPUBBPaKiiMIGEBgNVHR8EfTB7MHmgd6B1hnNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBJZGVudGl0eSUyMFZlcmlmaWNh
# dGlvbiUyMFJvb3QlMjBDZXJ0aWZpY2F0ZSUyMEF1dGhvcml0eSUyMDIwMjAuY3Js
# MIGUBggrBgEFBQcBAQSBhzCBhDCBgQYIKwYBBQUHMAKGdWh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSWRlbnRpdHklMjBW
# ZXJpZmljYXRpb24lMjBSb290JTIwQ2VydGlmaWNhdGUlMjBBdXRob3JpdHklMjAy
# MDIwLmNydDANBgkqhkiG9w0BAQwFAAOCAgEAX4h2x35ttVoVdedMeGj6TuHYRJkl
# FaW4sTQ5r+k77iB79cSLNe+GzRjv4pVjJviceW6AF6ycWoEYR0LYhaa0ozJLU5Yi
# +LCmcrdovkl53DNt4EXs87KDogYb9eGEndSpZ5ZM74LNvVzY0/nPISHz0Xva71Qj
# D4h+8z2XMOZzY7YQ0Psw+etyNZ1CesufU211rLslLKsO8F2aBs2cIo1k+aHOhrw9
# xw6JCWONNboZ497mwYW5EfN0W3zL5s3ad4Xtm7yFM7Ujrhc0aqy3xL7D5FR2J7x9
# cLWMq7eb0oYioXhqV2tgFqbKHeDick+P8tHYIFovIP7YG4ZkJWag1H91KlELGWi3
# SLv10o4KGag42pswjybTi4toQcC/irAodDW8HNtX+cbz0sMptFJK+KObAnDFHEsu
# kxD+7jFfEV9Hh/+CSxKRsmnuiovCWIOb+H7DRon9TlxydiFhvu88o0w35JkNbJxT
# k4MhF/KgaXn0GxdH8elEa2Imq45gaa8D+mTm8LWVydt4ytxYP/bqjN49D9NZ81co
# E6aQWm88TwIf4R4YZbOpMKN0CyejaPNN41LGXHeCUMYmBx3PkP8ADHD1J2Cr/6tj
# uOOCztfp+o9Nc+ZoIAkpUcA/X2gSMkgHAPUvIdtoSAHEUKiBhI6JQivRepyvWcl+
# JYbYbBh7pmgAXVswggefMIIFh6ADAgECAhMzAAAAWXzacemNXvXAAAAAAABZMA0G
# CSqGSIb3DQEBDAUAMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMgUlNBIFRpbWVz
# dGFtcGluZyBDQSAyMDIwMB4XDTI2MDEwODE4NTkwMVoXDTI3MDEwNzE4NTkwMVow
# geMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsT
# JE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMe
# blNoaWVsZCBUU1MgRVNOOjdCMUEtMDVFMC1EOTQ3MTUwMwYDVQQDEyxNaWNyb3Nv
# ZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1dGhvcml0eTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBAKYu5/40eEX+hT+5jFa146bid3dA4LnXYntv
# kP3CGw4LGARFhnvLMSJ/VtsubzDaeFnm7yb2KSM70WmHQprdCVqpvUH7l0uB4jNw
# 7urLoAR9kKHLE0VlMlDStDSxUBI3qwsdrjvdmvV0k+9/njuDEiSlzJTf7Dowd1K3
# bO4beRyaFhR+Y8tymECOqlOAffYrG2wZdVM51+QSBSe+PEykr8C6OnnqSipuF8fZ
# vCb6/huk0Zm6ZwsaixSHIAT2IEGvS7c63Im8jV3a8R0K6i2yiw0NNlnTSpwy/Zfv
# 7iwsLBwhfbjBTn+XOl6mPzDXQQ3V+SRP9xXbGKOsBTxzGid7aKAHw3o4Ahl9UGWL
# H9kNP3VUokE6JYkjlfpuUGZ6gQyqDewfxD4VoYIlopt4HZ0xQvqajuJx+cr8LR/I
# Z56gLLmwyMzde5+vtjBoilry/gSZwVGwgkvkIgpKPBQHGsSB0y3szr7Y7wEb6v0y
# Zal1XUvWnnz3inTaSWsCFrLPVwVmXy3ncY5/d25VpOkht+m697GWNbvsNOhAOHRa
# ftE9j/hhkoM6RsyJfBLnhqMcA/wcavf5oj5NeyRQdGZeLKcls9csKS3sBUzPidxx
# 2iiNH9CPaDq/bLJEOXasYohXMnRinu+fUk81s8VO7DQSF6ffn5oqSHoV8lf1Ax6u
# +kdShb8BAgMBAAGjggHLMIIBxzAdBgNVHQ4EFgQUj5bnC18D0vlnSRhCOiODGGuX
# NnYwHwYDVR0jBBgwFoAUa2koOjUvSGNAz3vYr0npPtk92yEwbAYDVR0fBGUwYzBh
# oF+gXYZbaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9z
# b2Z0JTIwUHVibGljJTIwUlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNy
# bDB5BggrBgEFBQcBAQRtMGswaQYIKwYBBQUHMAKGXWh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwUHVibGljJTIwUlNBJTIw
# VGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNydDAMBgNVHRMBAf8EAjAAMBYGA1Ud
# JQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDBmBgNVHSAEXzBdMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQCMA0GCSqG
# SIb3DQEBDAUAA4ICAQBEMhzC/ZcjpG/zURE7z2Yp5vrUxUjsE5Xa3t/2RGvESwvb
# msk3bLHhSFAajgo2XQ8xoGDP3sUhKCLPeICSbkVv6V8sSp8fJ8Jos6yrawf2YVis
# 8tcV+OO7U9S6JGPQzpmPncfzQc4ne1fqZ4+HiKabIDEoFdddQT2Egkk9fzxCY/EZ
# 52avJ27dSfrI/IDmyn9V10O3iQpg2F+C9vNTrk7nVgoDoHa9+Q3pYr0IHGnSmt5i
# rgGT436zo5WnXP8FxMhswH1aiyiSZiVzhor10C9C52cP3C8/PEoMKUXstLjoPO0T
# MkeW/1Fr186KXD45QRgBo0xImgtWTdzWFnlD+p7+iDBIuSrNcRXDRYuq/aYZaDhW
# SI0SYdPIWVh5XvXuWA31a8oQ0SO+oPa3Nk80k0864wiiyJ1KsbSnaaefg9vspegh
# rpY8ljCwxfCUtx5HQRNgAJOI8IKACK4d014Mk0hlRO0lQVRHegqIg29K6Xqkc360
# W2ZJGUcstlKokkVj6KAHjGyrLRPzepYfiZUJq4gXyxbpvKb1XJ2FN2682aUoNXo9
# RyRK1ch0f66k6+yj88kzvuC7+vJWtNDs/UpIM6Hhm0kU64JUJ7MMEQcAc7kpft7G
# m7YeRK+oKgqUgYXCfmzbX8nJXJZnPa8ADWVsIqsuNAxCI0CZXkULofqo5Be6zzGC
# A9QwggPQAgEBMHgwYTELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFB1YmxpYyBSU0EgVGltZXN0
# YW1waW5nIENBIDIwMjACEzMAAABZfNpx6Y1e9cAAAAAAAFkwDQYJYIZIAWUDBAIB
# BQCgggEtMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQgYx+FgoT2Fd7up9+dl+Oa8vBBIoEK4Za8rzTwHkDxvGEwgd0GCyqGSIb3DQEJ
# EAIvMYHNMIHKMIHHMIGgBCDLRbqx24bpscXEJ+Hjj9xrcUVw7R8OyyMfSB2YGK3+
# vDB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lc3RhbXBp
# bmcgQ0EgMjAyMAITMwAAAFl82nHpjV71wAAAAAAAWTAiBCCaD+AcIbcE8gvQ9SYi
# bUU49UK4bkXRYoo2okWo32SSxjANBgkqhkiG9w0BAQsFAASCAgBZkVh1OPKRUBBr
# 8mHjPz+xEz4QIT6UvaFJUcdjVTdlVxvkgwbyvkNs8WsZKsraZDRrYViXqE7gTsLf
# /XUNOnMlunG0x+M4bCwpAKGG3NshrpaCN8BY27sHyOZd6CqTnLKNHgm+GScShpxN
# vIfm33q5P54Nrvl1EbketPVPlKLuwotGhYcTWC93JfMjE9DVAPstghA+nhMjYabI
# e7+VZHslh8ejlF5xnmuqQ6bor7K4PDC6i9V4zILSZB/3t8NozVkUBwRMTnxiCoAv
# twCgLvIYigwh5rQyGreiblaBoelumQOnXdU/Uy1XHOl1ZDZhbAdfrkeo4smbKQlg
# +/96jlQgkYtb54OfnqCCzsYAmKogYGt/AmSDUbc6bUO6QM/xr1ccje28/WBv58PH
# LFn4TrEvDFce2cH+blRElvakzWTCExhaDyvymJNou7zruH5QV1O7kEIW5B9O/eSB
# cKg/PEwHiNuMLCbbOjJhonqz+x5eGuQOUJubFDRhDCCsNlReXUy+CN0RW8mt1J8z
# 0lfP9NUuBJQ3ajICZBV2lBVXGk4EuEQd7z2nNMFXt7/sEFtO+GznWKb1PEy4K3aV
# hLUedgey5/EMQXGc8pVvKB3VeG4KbAOgI5VJI5Usd/j/dDLHvbNUATeaFxAzdbbY
# 8HrXdUHWTIXAkQAaFFJsZfX7OyT8GQ==
# SIG # End signature block
