<#
.SYNOPSIS
    Registers the .grdp file type with Windows for GRIPS Direct Print.

.DESCRIPTION
    This script registers the .grdp file extension with the Windows registry.
    It associates .grdp files with the Print-GRDPFile.exe handler and sets
    appropriate icons and display names for all users on the system.

.PARAMETER ExePath
    Full path to Print-GRDPFile.exe.
    Defaults to the fully resolved path of '..\Print-GRDPFile.exe' relative
    to this script's location. The grips.ico icon is expected in the same
    directory as the executable.

.NOTES
    - Must be run with Administrator privileges
    - Registers file type for all users (HKEY_CLASSES_ROOT)

.EXAMPLE
    .\Register-GRDPFileType.ps1

.EXAMPLE
    .\Register-GRDPFileType.ps1 -ExePath 'C:\MyPath\Print-GRDPFile.exe'
#>

# Requires Administrator privileges
#Requires -RunAsAdministrator

param(
    [string]$ExePath = ""
)

$ErrorActionPreference = "Stop"

# Configuration
$fileExtension = ".grdp"
$progId = "GRIPS.DirectPrint.Archive"
$friendlyTypeName = "GRIPS Direct Print Archive"
$openWithDisplayName = "GRIPS Direct Print"
if ([string]::IsNullOrEmpty($ExePath)) {
    $ExePath = Join-Path (Split-Path -Parent $PSScriptRoot) "Print-GRDPFile.exe"
} else {
    $ExePath = (Resolve-Path $ExePath).Path
}
$exePath = $ExePath
$iconPath = Join-Path (Split-Path $exePath -Parent) 'grips.ico'
$mimeType = "application/x-grdp-archive"

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "GRIPS Direct Print File Type Registration" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# Check if executable file exists
if (-not (Test-Path $exePath)) {
    Write-Host "WARNING: Executable not found at: $exePath" -ForegroundColor Yellow
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
Write-Host "  Executable: $exePath" -ForegroundColor Gray
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
    # Call the executable directly with the file as an argument
    $commandValue = '"' + $exePath + '" "%1"'
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
    Write-Host "  - Execute: $exePath" -ForegroundColor Gray
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD92LXNJKPqT94X
# 58LrJVeUYIEoeMpujzEBuryNaKLRlaCCI9YwggXMMIIDtKADAgECAhBUmNLR1FsZ
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
# 03u4aUoqlmZpxJTG9F9urJh4iIAGXKKy7aIwggdaMIIFQqADAgECAhMzAAAABJZQ
# S9Lb7suIAAAAAAAEMA0GCSqGSIb3DQEBDAUAMGMxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xNDAyBgNVBAMTK01pY3Jvc29mdCBJ
# RCBWZXJpZmllZCBDb2RlIFNpZ25pbmcgUENBIDIwMjEwHhcNMjEwNDEzMTczMTUy
# WhcNMjYwNDEzMTczMTUyWjBaMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSswKQYDVQQDEyJNaWNyb3NvZnQgSUQgVmVyaWZpZWQg
# Q1MgQU9DIENBIDAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA4c6g
# 6DOiY6bAOwCPbBlQF2tjo3ckUZuab5ZorMnRp4rOmwZDiTbIpzFkZ/k8k4ivBJV1
# w5/b/oykI+eXAqaaxMdyAO0ModnEW7InfQ+rTkykEzHxRbCNg6KDsTnYc/YdL7II
# iJli8k51upaHLL7CYm9YNc0SFYvlaFj2O0HjO9y/NRmcWNjamZOlRjxW2cWgUsUd
# azSHgRCek87V2bM/17b+o8WXUW91IpggRasmiZ65WEFHXKbyhm2LbhBK6ZWmQoFe
# E+GWrKWCGK/q/4RiTaMNhHXWvWv+//I58UtOxVi3DaK1fQ6YLyIIGHzD4CmtcrGi
# vxupq/crrHunGNB7//Qmul2ZP9HcOmY/aptgUnwT+20g/A37iDfuuVw6yS2Lo0/k
# p/jb+J8vE4FMqIiwxGByL482PMVBC3qd/NbFQa8Mmj6ensU+HEqv9ar+AbcKwumb
# ZqJJKmQrGaSNdWfk2NodgcWOmq7jyhbxwZOjnLj0/bwnsUNcNAe09v+qiozyQQes
# 8A3UXPcRQb8G+c0yaO2ICifWTK7ySuyUJ88k1mtN22CNftbjitiAeafoZ9Vmhn5R
# fb+S/K5arVvTcLukt5PdTDQxl557EIE6A+6XFBpdsjOzkLzdEh7ELk8PVPMjQfPC
# gKtJ84c17fd2C9+pxF1lEQUFXY/YtCL+Nms9cWUCAwEAAaOCAg4wggIKMA4GA1Ud
# DwEB/wQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUJEWZoXeQKnzD
# yoOwbmQWhCr4LGcwVAYDVR0gBE0wSzBJBgRVHSAAMEEwPwYIKwYBBQUHAgEWM2h0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0
# bTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB8GA1UdIwQYMBaAFNlBKbAPD2Ns72nX9c0pnqRIajDmMHAGA1UdHwRpMGcwZaBj
# oGGGX2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29m
# dCUyMElEJTIwVmVyaWZpZWQlMjBDb2RlJTIwU2lnbmluZyUyMFBDQSUyMDIwMjEu
# Y3JsMIGuBggrBgEFBQcBAQSBoTCBnjBtBggrBgEFBQcwAoZhaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBJRCUyMFZlcmlm
# aWVkJTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDIxLmNydDAtBggrBgEFBQcw
# AYYhaHR0cDovL29uZW9jc3AubWljcm9zb2Z0LmNvbS9vY3NwMA0GCSqGSIb3DQEB
# DAUAA4ICAQBnLThdlbMNIokdKtzSa8io+pEO95Cc3VOyY/hQsIIcdMyk2hJOzLt/
# M1WXfQyElDk/QtyLzX63TdOb5J+nO8t0pzzwi7ZYvMiNqKvAQO50sMOJn3T3hCPp
# pxNNhoGFVxz2UyiQ4b2vOrcsLK9TOEFXWbUMJObR9PM0wZsABIhu4k6VVLxEDe0G
# SeQX/ZE7PHfTg44Luft4IKqYmnv1Cuosp3glFYsVegLnMWZUZ8UtO9F8QCiAouJY
# hL5OlCksgDb9ve/HQhLFnelfg6dQubIFsqB9IlConYKJZ/HaMZvYtA7y9EORK4cx
# lvTetCXAHayiSXH0ueE/T92wVG0csv5VdUyj6yVrm22vlKYAkXINKvDOB8+s4h+T
# gShlUa2ACu2FWn7JzlTSbpk0IE8REuYmkuyE/BTkk93WDMx7PwLnn4J+5fkvbjjQ
# 08OewfpMhh8SuPdQKqmZ40I4W2UyJKMMTbet16JFimSqDChgnCB6lwlpe0gfbo97
# U7prpbfBKp6B2k2f7Y+TjWrQYN+OdcPOyQAdxGGPBwJSaJG3ohdklCxgAJ5anCxe
# Yl7SjQ5Eua6atjIeVhN0KfPLFPpYz5CQU+JC2H79x4d/O6YOFR9aYe54/CGup7dR
# UIfLSv1/j0DPc6Elf3YyWxloWj8yeY3kHrZFaAlRMwhAXyPQ3rEX9zCCB38wggVn
# oAMCAQICEzMACIf1WPU0g4NtsvQAAAAIh/UwDQYJKoZIhvcNAQEMBQAwWjELMAkG
# A1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjErMCkGA1UE
# AxMiTWljcm9zb2Z0IElEIFZlcmlmaWVkIENTIEFPQyBDQSAwMjAeFw0yNjAzMjcx
# MDI0NTlaFw0yNjAzMzAxMDI0NTlaMIH8MRMwEQYDVQQREwo0NDMxNi0wMDAxMQsw
# CQYDVQQGEwJVUzENMAsGA1UECBMET2hpbzEOMAwGA1UEBxMFQWtyb24xGzAZBgNV
# BAkTEjIwMCBJbm5vdmF0aW9uIFdheTFNMEsGA1UECh5EAFQAaABlACAARwBvAG8A
# ZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBtAHAA
# YQBuAHkxTTBLBgNVBAMeRABUAGgAZQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQAaQBy
# AGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5MIIBojANBgkqhkiG
# 9w0BAQEFAAOCAY8AMIIBigKCAYEAs7Q1LdjcGXM18/F/vh6iPUMD+grySFti1iCg
# 2epcN/43iDbZIauutDVRcLNHCFDR5CmagPYGK7fsGsiKNZSx06haol/yNwrZ0pu8
# gpV+VidB82qQpnQXkEVNR3g+ms1JsF9Q5kdHSfgDroEoWW3fseLCnVb6seWJm/Eb
# Riho/SzbQ6EDEPyxRwdU0G7ffpwMRtGEd+U23YXaqXt9j1fDStqbJmEIxdXXNLyZ
# L2yuTKgsQDDc19JiK399cRRJuwCps1awUvCE6JJatUGGSLs9frTz+W5O67PZmpAg
# saYcgGaUNOsiouTwbgG1LuWifCI8GnAWM/pWctIV5hUxrE+BkticQwvHppQREuRs
# ++98TI/LTtw0310hhulqv6dikLW9b7MdK0t3eWTdTAXFoz3Gt2n29S3tCS5L5CPh
# W/5/pt5QRhblSteDKX7oor/TFx2A0J53FH23Zmr1y9EgFiatwyhILy1WitDS7yO5
# SFtnDsn1z97/EvxjHqyTh4xYdckzAgMBAAGjggIZMIICFTAMBgNVHRMBAf8EAjAA
# MA4GA1UdDwEB/wQEAwIHgDA8BgNVHSUENTAzBgorBgEEAYI3YQEABggrBgEFBQcD
# AwYbKwYBBAGCN2GDi6zmXoHCg+Bng6e09AGZrMsTMB0GA1UdDgQWBBQpOBumsczA
# 97i99dJbSK01XZ7jODAfBgNVHSMEGDAWgBQkRZmhd5AqfMPKg7BuZBaEKvgsZzBn
# BgNVHR8EYDBeMFygWqBYhlZodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBBT0MlMjBDQSUy
# MDAyLmNybDCBpQYIKwYBBQUHAQEEgZgwgZUwZAYIKwYBBQUHMAKGWGh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSUQlMjBW
# ZXJpZmllZCUyMENTJTIwQU9DJTIwQ0ElMjAwMi5jcnQwLQYIKwYBBQUHMAGGIWh0
# dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDBmBgNVHSAEXzBdMFEGDCsG
# AQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQBMA0GCSqGSIb3
# DQEBDAUAA4ICAQCGveutZrHCrsP3Y7Cdqb40zR4N4TREZvUKnpnt3aCHaEJ6a2aM
# TvxxsJr96QHQyKtAcRJpDB+PXY/RzBcDxpGBFYzUmecu2B4zJv/4cUISItYXSuWs
# WL3ypLbmxwlEzDxYZeRRc+RnerRcmJUHTAr4ZrFIn4UqYH5JSt8EXf1mXu8ioyZ9
# /VxIDWooVa0fMU25svx/YYsXwEUkrYoG7Vl//54kkoYPJAYBod1buY915MZXBIJC
# f5yNQbzy8K2Pnw8zQz8mkDkTWNvqGMR8por1cw5DczEawagyh2qKGBMNLc16g0g9
# prCbl6kVKahftjgxFQK4EHoyVGP9WWdN7Eq21CRHVXsjjPOaxPbQBH10DFfmOFVX
# /b8Wf4COfzVYMji4Hkgfy1v/MMHXNmm2UilAsZYLq2+P3XnywDUsk8gcGN2m8oid
# pcsmZkLMeUCUJyH9mm4OeM0yTSxEFo2LRMil97JN89862ig3Zjk4PR6lEXElUWww
# qbFv2QxzimaygSxuFGSDH8Kb3slBCs0+dw6onVUQCe7wV6DdgQ84TKk4RDSU666C
# /VW0En0r/AuR+x54/ZrY3fLkXyhzWFgxmWWjbjoO0s1YTBEDcMRjudaaquwRZ3yq
# apL9jsNj7rWai1hFmjrN8dsDJXMUdS3CvVT+IdTq5IFBPCPsr/NEOFAyTTCCB38w
# ggVnoAMCAQICEzMACIf1WPU0g4NtsvQAAAAIh/UwDQYJKoZIhvcNAQEMBQAwWjEL
# MAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjErMCkG
# A1UEAxMiTWljcm9zb2Z0IElEIFZlcmlmaWVkIENTIEFPQyBDQSAwMjAeFw0yNjAz
# MjcxMDI0NTlaFw0yNjAzMzAxMDI0NTlaMIH8MRMwEQYDVQQREwo0NDMxNi0wMDAx
# MQswCQYDVQQGEwJVUzENMAsGA1UECBMET2hpbzEOMAwGA1UEBxMFQWtyb24xGzAZ
# BgNVBAkTEjIwMCBJbm5vdmF0aW9uIFdheTFNMEsGA1UECh5EAFQAaABlACAARwBv
# AG8AZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBt
# AHAAYQBuAHkxTTBLBgNVBAMeRABUAGgAZQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQA
# aQByAGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5MIIBojANBgkq
# hkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAs7Q1LdjcGXM18/F/vh6iPUMD+grySFti
# 1iCg2epcN/43iDbZIauutDVRcLNHCFDR5CmagPYGK7fsGsiKNZSx06haol/yNwrZ
# 0pu8gpV+VidB82qQpnQXkEVNR3g+ms1JsF9Q5kdHSfgDroEoWW3fseLCnVb6seWJ
# m/EbRiho/SzbQ6EDEPyxRwdU0G7ffpwMRtGEd+U23YXaqXt9j1fDStqbJmEIxdXX
# NLyZL2yuTKgsQDDc19JiK399cRRJuwCps1awUvCE6JJatUGGSLs9frTz+W5O67PZ
# mpAgsaYcgGaUNOsiouTwbgG1LuWifCI8GnAWM/pWctIV5hUxrE+BkticQwvHppQR
# EuRs++98TI/LTtw0310hhulqv6dikLW9b7MdK0t3eWTdTAXFoz3Gt2n29S3tCS5L
# 5CPhW/5/pt5QRhblSteDKX7oor/TFx2A0J53FH23Zmr1y9EgFiatwyhILy1WitDS
# 7yO5SFtnDsn1z97/EvxjHqyTh4xYdckzAgMBAAGjggIZMIICFTAMBgNVHRMBAf8E
# AjAAMA4GA1UdDwEB/wQEAwIHgDA8BgNVHSUENTAzBgorBgEEAYI3YQEABggrBgEF
# BQcDAwYbKwYBBAGCN2GDi6zmXoHCg+Bng6e09AGZrMsTMB0GA1UdDgQWBBQpOBum
# sczA97i99dJbSK01XZ7jODAfBgNVHSMEGDAWgBQkRZmhd5AqfMPKg7BuZBaEKvgs
# ZzBnBgNVHR8EYDBeMFygWqBYhlZodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBBT0MlMjBD
# QSUyMDAyLmNybDCBpQYIKwYBBQUHAQEEgZgwgZUwZAYIKwYBBQUHMAKGWGh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSUQl
# MjBWZXJpZmllZCUyMENTJTIwQU9DJTIwQ0ElMjAwMi5jcnQwLQYIKwYBBQUHMAGG
# IWh0dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDBmBgNVHSAEXzBdMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQBMA0GCSqG
# SIb3DQEBDAUAA4ICAQCGveutZrHCrsP3Y7Cdqb40zR4N4TREZvUKnpnt3aCHaEJ6
# a2aMTvxxsJr96QHQyKtAcRJpDB+PXY/RzBcDxpGBFYzUmecu2B4zJv/4cUISItYX
# SuWsWL3ypLbmxwlEzDxYZeRRc+RnerRcmJUHTAr4ZrFIn4UqYH5JSt8EXf1mXu8i
# oyZ9/VxIDWooVa0fMU25svx/YYsXwEUkrYoG7Vl//54kkoYPJAYBod1buY915MZX
# BIJCf5yNQbzy8K2Pnw8zQz8mkDkTWNvqGMR8por1cw5DczEawagyh2qKGBMNLc16
# g0g9prCbl6kVKahftjgxFQK4EHoyVGP9WWdN7Eq21CRHVXsjjPOaxPbQBH10DFfm
# OFVX/b8Wf4COfzVYMji4Hkgfy1v/MMHXNmm2UilAsZYLq2+P3XnywDUsk8gcGN2m
# 8oidpcsmZkLMeUCUJyH9mm4OeM0yTSxEFo2LRMil97JN89862ig3Zjk4PR6lEXEl
# UWwwqbFv2QxzimaygSxuFGSDH8Kb3slBCs0+dw6onVUQCe7wV6DdgQ84TKk4RDSU
# 666C/VW0En0r/AuR+x54/ZrY3fLkXyhzWFgxmWWjbjoO0s1YTBEDcMRjudaaquwR
# Z3yqapL9jsNj7rWai1hFmjrN8dsDJXMUdS3CvVT+IdTq5IFBPCPsr/NEOFAyTTCC
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
# BgNVBAMTIk1pY3Jvc29mdCBJRCBWZXJpZmllZCBDUyBBT0MgQ0EgMDICEzMACIf1
# WPU0g4NtsvQAAAAIh/UwDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3AgEMMQIw
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQgVVoz
# rSzZOCvzoyXhyjPz0eOzOHklaMF62caumTkIPAEwDQYJKoZIhvcNAQEBBQAEggGA
# MBteRwla0R/17gBuGmrplyNhLict313/MeDwd2KHNw7RKjJOksKziuXZHTa/AZUc
# WRTu3do8T0Rd1bVsrSvOi8o2UWXOZPNXrQNyKSfCOAfT9k46CkDS6b4yYznbdzcY
# 7Yo7dvcl1vQaTe5jL9MhV0/hxBDV8lXD4muIm4RNQQeCj0zAACttLrM7vxW1YWFW
# 2srtP2d/xMe7+VfheytNNrQZ+b7YW0i7QK7k2UaaMyLNFyb3GnO6v7Egr0a3rL+a
# h6gSN0cNeOihIiaPpXANaqf+s0mhPfI5Q3oYdv5NuSn77w4MRGzLH8VcCknuswmW
# CwsxwuQP3/kbC0bpcSwYcbRc5baRPwkzHehHmVX17GlWaFE5uODAYm9D72feJ4Pg
# fD1E/YsU/gohYW75EOH8eQa4dqTn9kJKEPOI+/Tv3NtyTzus8Gg0h54cEILza3Jq
# po4kjC5hkdA73MtN2rDEQe2mkujwVkrk0hgXB8kIQPnpFz8MPAfYCDp6gJcp1lam
# oYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIUizCCFIcC
# AQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSCAVUwggFR
# AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEID/yDcs16G9+Cm+losun
# pEnaKdVLkWZGn57iE0Y6wcWQAgZpxmf0X04YEzIwMjYwMzI3MTQwMDA5LjQxNlow
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
# IgQgjpesGKX4BZDQqcKbMkBuo9rkY017r2Dhul173DNT0CYwgd0GCyqGSIb3DQEJ
# EAIvMYHNMIHKMIHHMIGgBCDLRbqx24bpscXEJ+Hjj9xrcUVw7R8OyyMfSB2YGK3+
# vDB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lc3RhbXBp
# bmcgQ0EgMjAyMAITMwAAAFl82nHpjV71wAAAAAAAWTAiBCDFK+HlFh8xmJI1IZRw
# Xua6jvrh0DP/h0DNGeKxjSuf5jANBgkqhkiG9w0BAQsFAASCAgCjF3JKdwKOAM75
# BWUWil+5gv/nsKr2bGl2FIWrdhvsTBsN0lr9zOpMaBxO7GRbJYeErqWVMT34ZRAT
# OOImrBx/musPPsWicFuaL9Qwh6kds0/WgNbENtVb6Q+CBPv1Kbvzsomu6pj6XgCP
# +ukxZtYYMtRh1w1nFJapBKm9KOSelDUI75j7NAPX3GohYEagQkzPRUT2Y+yjsTks
# 0S6wQTUUO/IfjiPhVi3oik8t5S9jiwzW6ePG1VMO7bDWkeVLsDzM4AN+af1XjA5Z
# SAWouUaWO60LeykrOJDuL82XyXNAOMmhAHleXjV2naYDjraVjUWgBQTj9GSsasME
# 76zAyVqPIke52KSy0Bd5l68Bu4cZiTVsQ8lJs6CXaaMOxWW8JXywrp2X6MJdEqvb
# cGp4WaacYz4TVhYT9pDFkZVuLDg+Jv3MCeKBE01xGcXDOWYSKk+PR+NXPneSaud3
# dt2fgg+1pvtUyiT2C5OQ1drhvuU1K7MBAk0O2+i0PJ1xuLwxduzYmqnbCbMHTnOc
# hqdNeSYCuOnGCH3VcioNaD+X0mZGm60WYaEku2lT5EsTOw3lTEy9yWfWkejodI/J
# lDVGFg2OnIi5uOMWAXhMUZDhJufETfu5JvxH5tOKriGhIq0p2mDs+9jtgDkk9tjn
# T9x1qxfu7rJYNhK0PhP276VO5cYIdA==
# SIG # End signature block
