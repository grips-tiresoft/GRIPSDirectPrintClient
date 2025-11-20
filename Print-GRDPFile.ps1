param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    [string]$configFile = "",
    [string]$userConfigFile = ""
)

# Get the full path of the directory containing the script
$ScriptPath = $PSScriptRoot

if ($configFile -eq "") { $configFile = "$ScriptPath\config.json" }
$global:configFile = $configFile

if ($userConfigFile -eq "") { $userConfigFile = "$PSScriptRoot\userconfig.json" }
$global:userConfigFile = $userConfigFile

# Function to parse key=value pairs from a text file into a hashtable
function Get-Options {
    param([string]$FilePath)
    $options = @{}
    Get-Content $FilePath | ForEach-Object {
        if ($_ -match '^\s*([^=]+)\s*=\s*(.+)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $options[$key] = $value
        }
    }
    return $options
}

function Update-Check {
    Write-Output "Checking for updates..."
    
    if ($global:config.UsePrereleaseVersion) {
        Write-Output "Checking for latest release (including prereleases)..."
        # Get all releases (sorted by date, newest first)
        $AllReleases = Invoke-RestMethod -Uri ($releaseApiUrl -replace '/latest$', '') -Method Get
        $LatestRelease = $AllReleases | Select-Object -First 1
    }
    else {
        Write-Output "Checking for latest stable release only..."
        $LatestRelease = Invoke-RestMethod -Uri $releaseApiUrl -Method Get
    }
    $releaseVersion = $LatestRelease.tag_name.TrimStart('v')
    Get-ScriptVersion

    # Compare versions
    if ([version]$releaseVersion -gt [version]$global:currentVersion) {
        # The latest version is greater than the current version
        $TempZipFile = [System.IO.Path]::GetTempFileName() + ".zip"
        $TempExtractPath = [System.IO.Path]::GetTempPath() + [System.Guid]::NewGuid().ToString()

        # Get the URL of the source code zip
        $downloadUrl = $LatestRelease.zipball_url

        # Download the ZIP file containing the new script version and other files
        Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZipFile

        # Extract the ZIP file to a temporary directory
        Expand-Archive -Path $TempZipFile -DestinationPath $TempExtractPath

        # Find the sub-folder in the extracted directory
        $extractedSubFolder = Get-ChildItem -Path $TempExtractPath | Where-Object { $_.PSIsContainer } | Select-Object -First 1

        # Ensure the file exists before writing to it
        if (-not (Test-Path -Path $updateSignalFile)) {
            New-Item -Path $updateSignalFile -ItemType File -Force
        }
    
        # Clean up temporary files
        Remove-Item -Path $TempZipFile -Force

        # Signal the main script that the update is ready
        Set-Content -Path $updateSignalFile -Value "$($extractedSubFolder.FullName)"
    }
    else {
        Write-Output "No update required. Current version ($global:currentVersion) is up to date."
    }
}

# Function to perform the update
function Update-Release {
    # Ensure the update signal file exists before trying to read it
    if (-not (Test-Path -Path $updateSignalFile)) {
        Write-Error "Update signal file not found at: $updateSignalFile"
        return
    }
    
    # Read the path of the extracted folder
    $extractedSubFolder = Get-Content -Path $updateSignalFile
    
    if ([string]::IsNullOrWhiteSpace($extractedSubFolder)) {
        Write-Error "Update signal file is empty: $updateSignalFile"
        Remove-Item -Path $updateSignalFile -Force
        return
    }
	
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Updating release from $extractedSubFolder"

    # Backup the current script directory
    $backupScriptDirectory = "$ScriptPath.bak"
	
    Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Backup script folder: $backupScriptDirectory"
	
    if (Test-Path -Path $backupScriptDirectory) {
        Remove-Item -Path $backupScriptDirectory -Recurse -ErrorAction SilentlyContinue
    }
    
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Backing up current script folder from $ScriptPath to $backupScriptDirectory"
    # Copy the script directory to the backup directory
    #Copy-Item -Path $ScriptPath -Destination $backupScriptDirectory -Recurse -Force -Exclude '$Recycle.Bin'
    $robocopyCommand = @"
robocopy "$ScriptPath" "$backupScriptDirectory" /E /XD '`$Recycle.Bin'
"@
    Invoke-Expression $robocopyCommand
	
    # Copy the extracted files from the sub-folder to the destination directory
    $resolvedPath = Resolve-Path -Path $extractedSubFolder
	
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Copying new script folder from $resolvedPath to $ScriptPath"
    Copy-Item -Path "$resolvedPath\*" -Destination $ScriptPath -Recurse -Force

    Remove-Item -Path $updateSignalFile -Force

    Get-ScriptVersion
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Script updated to version $global:currentVersion."
}

function Get-Config {
    # Load configuration from JSON file
    $global:config = Get-Content $global:configFile -Encoding UTF8 | ConvertFrom-Json

    # Check if userconfig.json exists
    if (Test-Path -Path $global:userconfigFile -PathType Leaf) {
        # Load user configuration from userconfig.json
        $global:userConfig = Get-Content $global:userConfigFile -Encoding UTF8 | ConvertFrom-Json

        # Update or add keys from user configuration
        $global:userConfig.PSObject.Properties | ForEach-Object {
            $global:config | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value -Force
        }    
    }

    $config

    # Load language strings
    $global:LanguageStrings = Get-LanguageStrings
}

function Get-ScriptVersion {
    Get-Config -configFile $global:configFile -userConfigFile $global:userConfigFile

    $global:currentVersion = $global:config.Version.TrimStart('v')
    Write-Output "Script version: $global:currentVersion"
}

# Function to get the last update check time
function Get-LastUpdateCheckTime {
    if (Test-Path $lastUpdateCheckFile) {
        $content = Get-Content $lastUpdateCheckFile -ErrorAction SilentlyContinue
        if ($content -and $content.Trim() -ne "") {
            try {
                # Use Parse instead of TryParse, catch exceptions if invalid
                $parsedDate = [DateTime]::Parse($content)
                return $parsedDate
            }
            catch {
                # Parsing failed, return MinValue
                return [DateTime]::MinValue
            }
        }
    }
    return [DateTime]::MinValue
}

# Function to set the last update check time to now
function Set-LastUpdateCheckTime {
    $now = Get-Date
    Set-Content -Path $lastUpdateCheckFile -Value $now.ToString("o") # ISO 8601 format
}

# Function to start a new transcript with a timestamped filename
function Start-MyTranscript {
    param (
        [string]$Path = "$ScriptPath\Transcripts",
        [string]$Filename = "$ScriptNameWithoutExt"
    )

    if (-not (Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $transcriptPath = Join-Path -Path $Path -ChildPath "$($Filename)_$timestamp.Transcript.txt"
    Start-Transcript -Path $transcriptPath | Out-Null

    # Remove old transcript files
    $transcriptPath = Join-Path -Path $Path -ChildPath "$($Filename)_*.Transcript.txt"
    $transcripts = Get-ChildItem -Path $transcriptPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
    $transcripts | Remove-Item -Force

    return [datetime]::Now
}

# Return a unique filename by appending (1), (2), etc. if the file already exists
function Get-UniqueFileName {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        return $FilePath
    }
    
    $directory = Split-Path $FilePath -Parent
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    
    $counter = 1
    do {
        $newPath = Join-Path $directory "$filename ($counter)$extension"
        $counter++
    } while (Test-Path $newPath)
        
    return $newPath
}

# Function to check if a printer exists
function Test-PrinterExists {
    param([string]$PrinterName)
    try {
        $printer = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        return ($null -ne $printer)
    }
    catch {
        return $false
    }
}

# Function to load language strings
function Get-LanguageStrings {
    param([string]$LanguageFile = "$ScriptPath\languages.json")
    
    if (-not (Test-Path $LanguageFile)) {
        Write-Warning "Language file not found: $LanguageFile. Using default English strings."
        return $null
    }
    
    try {
        $allLanguages = Get-Content $LanguageFile -Encoding UTF8 | ConvertFrom-Json
        
        # Get OS culture
        $osCulture = [System.Globalization.CultureInfo]::CurrentUICulture.Name
        Write-Host "OS Language: $osCulture"
        
        # Try exact match first (e.g., en-US)
        if ($allLanguages.PSObject.Properties.Name -contains $osCulture) {
            Write-Host "Using language strings for: $osCulture"
            return $allLanguages.$osCulture
        }
        
        # Try language-only match (e.g., en from en-US)
        $languageOnly = $osCulture.Split('-')[0]
        $matchingLanguage = $allLanguages.PSObject.Properties.Name | Where-Object { $_ -like "$languageOnly-*" } | Select-Object -First 1
        
        if ($matchingLanguage) {
            Write-Host "Using language strings for: $matchingLanguage (matched from $languageOnly)"
            return $allLanguages.$matchingLanguage
        }
        
        # Fall back to en-US
        Write-Host "No matching language found. Using en-US as fallback."
        return $allLanguages.'en-US'
    }
    catch {
        Write-Error "Failed to load language file: $_"
        return $null
    }
}

# Function to show printer selection dialog
function Select-AlternativePrinter {
    param(
        [string]$MissingPrinterName,
        [PSCustomObject]$LanguageStrings
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Use default English strings if language strings not loaded
    if ($null -eq $LanguageStrings) {
        $LanguageStrings = [PSCustomObject]@{
            PrinterNotFound = "Printer '{0}' not found.`n`nSelect an alternative printer:"
            NoPrinters = "No printers available on this system."
            NoPrintersTitle = "No Printers"
            PrinterNotFoundTitle = "Printer Not Found"
            OK = "OK"
            Cancel = "Cancel"
        }
    }
    
    # Get list of available printers
    $printers = Get-Printer | Select-Object -ExpandProperty Name
    
    if ($printers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $LanguageStrings.NoPrinters,
            $LanguageStrings.NoPrintersTitle,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $null
    }
    
    # Create form for printer selection
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $LanguageStrings.PrinterNotFoundTitle
    $form.Size = New-Object System.Drawing.Size(400, 380)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    
    # Warning Label
    $warningLabel = New-Object System.Windows.Forms.Label
    $warningLabel.Location = New-Object System.Drawing.Point(10, 10)
    $warningLabel.Size = New-Object System.Drawing.Size(360, 60)
    $warningLabel.Text = $LanguageStrings.PrinterNotFound -f $MissingPrinterName
    $warningLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($warningLabel)
    
    # ListBox
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 80)
    $listBox.Size = New-Object System.Drawing.Size(360, 200)
    $listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    
    foreach ($printer in $printers) {
        [void]$listBox.Items.Add($printer)
    }
    
    if ($listBox.Items.Count -gt 0) {
        $listBox.SelectedIndex = 0
    }
    
    $form.Controls.Add($listBox)
    
    # OK Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(210, 310)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = $LanguageStrings.OK
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    # Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(295, 310)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = $LanguageStrings.Cancel
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    # Show dialog
    $dialogResult = $form.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and $listBox.SelectedItem) {
        return $listBox.SelectedItem
    }
    
    return $null
}

Get-Config -configFile $global:configFile -userConfigFile $global:userConfigFile

if (-not [System.IO.Path]::IsPathRooted($config.PDFPrinter_exe)) {
    $PDFPrinter_exe = "$ScriptPath\$($config.PDFPrinter_exe)"
}
else {
    $PDFPrinter_exe = $config.PDFPrinter_exe
}
$PDFPrinter_params = $config.PDFPrinter_params
$releaseApiUrl = $global:config.ReleaseApiUrl;
#$ReleaseCheckDelay = 600 # Delay between checking for new releases in seconds
$ReleaseCheckDelay = $global:config.ReleaseCheckDelay

# Get the filename of the script
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)
Start-MyTranscript 

# Main logic
try {
    if ($InputFile.ToLower().EndsWith(".grdp")) {
        # Create temp folder
        $tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempFolder | Out-Null

        try {
            # Extract the .grdp (zip) file
            $OldInputFile = $InputFile
            $InputFile = Join-Path $tempFolder ([System.IO.Path]::GetFileNameWithoutExtension($InputFile) + ".zip")
            Copy-Item $OldInputFile $InputFile
            Expand-Archive -Path $InputFile -DestinationPath $tempFolder -Force

            # Find printersettings.json
            $settingsFile = Join-Path $tempFolder "printsettings.json"
            if (-not (Test-Path $settingsFile)) {
                Write-Error "printersettings.json not found in archive."
                exit 1
            }

            $settings = Get-Content $settingsFile -Encoding UTF8 | ConvertFrom-Json

            foreach ($entry in $settings) {
                $filename = $entry.Filename
                $printer = $entry.Printer
                $outputBin = $entry.OutputBin
                $addArgs = $entry.AdditionalArgs

                $filePath = Join-Path $tempFolder $filename
                if (-not (Test-Path $filePath)) {
                    Write-Warning "File $filename not found in archive, skipping."
                    continue
                }

                if ($filePath.ToLower().EndsWith(".pdf")) {
                    # Check if printer exists
                    if (-not (Test-PrinterExists -PrinterName $printer)) {
                        Write-Warning "Printer '$printer' not found."
                        $alternativePrinter = Select-AlternativePrinter -MissingPrinterName $printer -LanguageStrings $global:LanguageStrings
                        
                        if ($null -eq $alternativePrinter) {
                            Write-Warning "No alternative printer selected. Skipping print job for $filePath"
                            continue
                        }
                        
                        Write-Host "Using alternative printer: $alternativePrinter"
                        $printer = $alternativePrinter
                    }
                    
                    # Construct paper source argument if OutputBin is specified
                    $paperSourceArg = if ([string]::IsNullOrEmpty($outputBin)) { "" } else { "bin={0}," -f $outputBin }

                    # Handle AdditionalArgs for -print-settings
                    if (-not [string]::IsNullOrEmpty($addArgs)) {
                        if ($addArgs.Contains("-print-settings") -and $PDFPrinter_params.Contains("-print-settings")) {
                            $addPrintArgs = $addArgs -split '\s+'
                            if ($addPrintArgs.Length -gt 1) {
                                $addArgs = $addPrintArgs[1].Trim('"')
                            }
                        }
                    }

                    $params = $PDFPrinter_params -f $printer, $filePath, $paperSourceArg, $addArgs

                    # Start printing
                    Write-Host "Printing $filePath to printer '$printer' with settings '$params'"
                    $proc = Start-Process -FilePath $PDFPrinter_exe -ArgumentList $params -PassThru

                    # Wait for process exit with timeout (e.g., 30 seconds)
                    if (-not $proc.WaitForExit(30000)) {
                        Write-Warning "Print process did not exit within 30 seconds, killing process."
                        try { $proc.Kill() } catch { Write-Warning "Failed to kill print process: $_" }
                    }
                    else {
                        Write-Host "Print job completed for $filePath"
                    }
                    continue
                }
                else {
                    # Open file with associated executable
                    $downloadsFolder = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
                    $uniqueFilePath = Get-UniqueFileName -FilePath (Join-Path -Path $downloadsFolder -ChildPath ([System.IO.Path]::GetFileName($filePath)))
                    Copy-Item -Path $filePath -Destination $uniqueFilePath
                    Write-Host "Opening signature file: $uniqueFilePath"
                    Start-Process -FilePath $uniqueFilePath
                    continue
                }
            }        
        }
        finally {
            # Clean up temp folder
            Remove-Item -Path $tempFolder -Recurse -Force
            
            # Remove old download files
            $downloadsFolder = [Environment]::GetFolderPath('UserProfile') + "\Downloads"

            # Remove old .eml files
            $downloadsPath = Join-Path -Path $downloadsFolder -ChildPath "NewEmail*.eml"
            $downloads = Get-ChildItem -Path $downloadsPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
            if ($null -ne $downloads) { 
                Write-Host "Removing old .eml files:" 
                Write-Host $downloads
                $downloads | Remove-Item -Force 
            }

            # Remove old .sig files
            $downloadsPath = Join-Path -Path $downloadsFolder -ChildPath "*.sig"
            $downloads = Get-ChildItem -Path $downloadsPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
            if ($null -ne $downloads) {
                Write-Host "Removing old .sig files:"
                Write-Host $downloads
                $downloads | Remove-Item -Force
            }

            # Remove old .grdp files
            $downloadsPath = Join-Path -Path $downloadsFolder -ChildPath "*.grdp"
            $downloads = Get-ChildItem -Path $downloadsPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
            if ($null -ne $downloads) {
                Write-Host "Removing old .grdp files:"
                Write-Host $downloads
                $downloads | Remove-Item -Force
            }
        }
    }
    else {
        # Normal PDF file - print to default printer
        Write-Host "Printing $InputFile to default printer"
        Start-Process -FilePath $PDFPrinter_exe -ArgumentList "-print-to-default", "`"$InputFile`"" -PassThru
    }
}
finally {
    # Define update signal file
    $updateSignalFile = "$ScriptPath\update_ready.txt"
    if (Test-Path -Path $updateSignalFile) {
        Update-Release
    }
    else {

        # Define a file to store the last update check timestamp
        $lastUpdateCheckFile = Join-Path -Path $ScriptPath -ChildPath "last_update_check.txt"

        # After printing completes, check if update check is needed
        $lastCheckTime = Get-LastUpdateCheckTime
        $now = Get-Date
        $elapsedSeconds = ($now - $lastCheckTime).TotalSeconds

        if ($elapsedSeconds -ge $ReleaseCheckDelay) {
            Set-LastUpdateCheckTime
            Write-Host "Time since last update check: $elapsedSeconds seconds. Checking for updates..."
            Update-Check
        }
        else {
            Write-Host "Last update check was $elapsedSeconds seconds ago. Skipping update check."
        }
    }
    Stop-Transcript
}

<# IF THE SCRIPT HAS BEEN CHANGED THEN IT WILL NEED RESIGNING:
.\CreateSignedScript.ps1 -Path .\Print-GRDPFile.ps1
#>

# SIG # Begin signature block
# MIIP/QYJKoZIhvcNAQcCoIIP7jCCD+oCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDypKCbUCZ8AevS
# q6iZE8cbcbGshICYX4tfdqiHEW1aWKCCDSgwggVTMIIEO6ADAgECAhMYAAAWfb+K
# c6CcOqMqAAAAABZ9MA0GCSqGSIb3DQEBCwUAMGIxLTArBgNVBAoTJFRoZSBHb29k
# eWVhciBUaXJlIGFuZCBSdWJiZXIgQ29tcGFueTExMC8GA1UEAxMoR29vZHllYXIg
# UHJvZHVjdGlvbiBHZW5lcmFsIFB1cnBvc2UgQ0EgMjAeFw0yNDAyMjgxMTM0Mjla
# Fw0yNjAyMjcxMTM0MjlaMIHEMQswCQYDVQQGEwJVUzELMAkGA1UECBMCT0gxDjAM
# BgNVBAcTBUFrcm9uMSswKQYDVQQKDCJUaGUgR29vZHllYXIgVGlyZSAmIFJ1YmJl
# ciBDb21wYW55MTUwMwYDVQQLDCxIb3N0ZWQgYnkgVGhlIEdvb2R5ZWFyIFRpcmUg
# JiBSdWJiZXIgQ29tcGFueTEOMAwGA1UECxMFR1JJUFMxJDAiBgNVBAMMG0dSSVBT
# X0NvZGVzaWduLmdvb2R5ZWFyLmNvbTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBAN22qYYRXYdgQ7aUyLfcOJg4D94BCxZ2o4/ZeNAYW5XaVMayJ3iE/ePq
# rtmN4qlLW3NZGxIAGRmAyB1Fbj217P1g1/6Urk058BOvTOylTRvvLTqRlLT+svJN
# IZWWEVXutBzdxyLNK/jBiJNA/jEsAn3eoYcwokOw3/AW5Q6typrUtJJdkLHNskE7
# O6z5g0JZ7jJKycPJa17jEe0ZQwigktnW41uTnGi3208RlXdYZejVYdCQ8sW+4Q8y
# f0HGEVvc0E1rycThhlxSN2a3unoqM3BTNSSotIeURePWo1Ke1eKmX9x2LLa8pbHg
# e1XiSRigiDUKG1zYeqB7CZf9u5FiFs0CAwEAAaOCAZ0wggGZMAsGA1UdDwQEAwIH
# gDAdBgNVHQ4EFgQURXgD76meYvLqQorZhHQde46d/g4wHwYDVR0jBBgwFoAU8d2+
# AB+LeGck24H2ygXFWVV0AxkwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3BraS5n
# b29keWVhci5jb20vR29vZHllYXIlMjBQcm9kdWN0aW9uJTIwR2VuZXJhbCUyMFB1
# cnBvc2UlMjBDQSUyMDIuY3JsMGoGCCsGAQUFBwEBBF4wXDBaBggrBgEFBQcwAoZO
# aHR0cDovL3BraS5nb29keWVhci5jb20vR29vZHllYXIlMjBQcm9kdWN0aW9uJTIw
# R2VuZXJhbCUyMFB1cnBvc2UlMjBDQSUyMDIuY3J0MAwGA1UdEwEB/wQCMAAwPQYJ
# KwYBBAGCNxUHBDAwLgYmKwYBBAGCNxUIhdK+QYSUwFmF6YEmhPTsUIbHtXYAhNOI
# HoL8sEwCAWQCAQ0wEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYJKwYBBAGCNxUKBA4w
# DDAKBggrBgEFBQcDAzANBgkqhkiG9w0BAQsFAAOCAQEAAaGkZcE0La+/rbafJMrb
# bI1zxYyoEgspBKR3ikjECie30EsvsasQqyih3CMNtJ7wwHKzs3dtIWg4/yz4I3EQ
# oEfmELwUcU2aQXS8JFvHVLwqL+eSgTuJkomDnnfopoxej3Z28Dmsfs/Gct9SYmc5
# gTHsZ6S2PWj9A6rR42uMAO9HhtJf2xKoMmSdyuDQEGfeTRWybPVUM7kjBQXTB7aC
# 41cRNSYa5ai0dv6fRcbEe/kkB9EheO6LP8k6jgR69oaYa5TA1HpTODyRbNI34tdd
# KdYEtjW3mlGxgQcdea8IqlVl/uRGPf3rb+U9ZvmSFlHSrDN0gvpkNXrqgotEXmHN
# 5jCCB80wgga1oAMCAQICExsAAAADnROFF+vViDwAAAAAAAMwDQYJKoZIhvcNAQEL
# BQAwVzEtMCsGA1UEChMkVGhlIEdvb2R5ZWFyIFRpcmUgYW5kIFJ1YmJlciBDb21w
# YW55MSYwJAYDVQQDEx1Hb29keWVhciBQcm9kdWN0aW9uIFJvb3QgQ0EgMjAeFw0x
# NzAzMDgxNjQwMzFaFw0zMDAzMDgxNjI0NTFaMGIxLTArBgNVBAoTJFRoZSBHb29k
# eWVhciBUaXJlIGFuZCBSdWJiZXIgQ29tcGFueTExMC8GA1UEAxMoR29vZHllYXIg
# UHJvZHVjdGlvbiBHZW5lcmFsIFB1cnBvc2UgQ0EgMjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAKt5jMVDPEy45K60OPKp0LZD9vnv3Qh7xttcA3nY0Szc
# 1HdH40yBtTtvpzM5lUfEs9nh9CMB/pqh5zorGk5Ltt7s+lGxWqn8p8C7YBE7t2un
# EW7gpksKK00/Oi73rFOYmiUusoTmmri0H1D6kz6w4LstIbPFFTMpfjrcFAPHepHx
# H4pxywYQ7jv6wobDgTtFGGkn+ClOZTRL6KDUAb/v124wQZVBVH76xJ2UFeE9Olug
# Xv/ELpSCRCvoaz6qNw4iV+RS5G+Uuy7PgsTPSewfN2TscXexS97VNk1v23r/wDKR
# p/nwH1n/d0th5ZNs3LIZjloVi8MtErPdl1VhV41U1osCAwEAAaOCBIUwggSBMA4G
# A1UdDwEB/wQEAwIBBjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU8d2+AB+L
# eGck24H2ygXFWVV0AxkwggM5BgNVHSAEggMwMIIDLDCCAygGDCsGAQQBgpgSAQEB
# ATCCAxYwggMSBggrBgEFBQcCAjCCAwQeggMAAFQAaABpAHMAIABDAGUAcgB0AGkA
# ZgBpAGMAYQB0AGkAbwBuACAAQQB1AHQAaABvAHIAaQB0AHkAIAAoAEMAQQApACAA
# aQBzACAAYQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQAaQByAGUAIAAmACAAUgB1AGIA
# YgBlAHIAIABDAG8AbQBwAGEAbgB5ACAAaQBuAHQAZQByAG4AYQBsACAAcgBlAHMA
# bwB1AHIAYwBlAC4AIAAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQBzACAAaQBzAHMA
# dQBlAGQAIABiAHkAIAB0AGgAaQBzACAAQwBBACAAYQByAGUAIABmAG8AcgAgAGkA
# bgB0AGUAcgBuAGEAbAAgAEcAbwBvAGQAeQBlAGEAcgAgAFQAaQByAGUAIAAmACAA
# UgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5ACAAdQBzAGUAIABvAG4AbAB5AC4A
# IAAgAEEAbgB5ACAAbgBvAG4ALQBHAG8AbwBkAHkAZQBhAHIAIABUAGkAcgBlACAA
# JgAgAFIAdQBiAGIAZQByACAAQwBvAG0AcABhAG4AeQAgAHAAYQByAHQAeQAgAHMA
# aABhAGwAbAAgAG4AbwB0ACAAcgBlAGwAeQAgAG8AbgAgAHQAaABpAHMAIABDAEEA
# IABmAG8AcgAgAGEAbgB5ACAAcgBlAGEAcwBvAG4ALgAgACAARgBvAHIAIABtAG8A
# cgBlACAAaQBuAGYAbwByAG0AYQB0AGkAbwBuACAAYQBiAG8AdQB0ACAAVABoAGUA
# IABHAG8AbwBkAHkAZQBhAHIAIABUAGkAcgBlACAAJgAgAFIAdQBiAGIAZQByACAA
# QwBvAG0AcABhAG4AeQAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAFAAbwBsAGkA
# YwB5ACwAIABwAGwAZQBhAHMAZQAgAGUAbQBhAGkAbAAgAHMAZQBjAHUAcgBpAHQA
# eQBAAGcAbwBvAGQAeQBlAGEAcgAuAGMAbwBtMBkGCSsGAQQBgjcUAgQMHgoAUwB1
# AGIAQwBBMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAUaVAMD1nLH3Jw
# W3ED/p+tvS32AZwwUgYDVR0fBEswSTBHoEWgQ4ZBaHR0cDovL3BraS5nb29keWVh
# ci5jb20vR29vZHllYXIlMjBQcm9kdWN0aW9uJTIwUm9vdCUyMENBJTIwMi5jcmww
# XQYIKwYBBQUHAQEEUTBPME0GCCsGAQUFBzAChkFodHRwOi8vcGtpLmdvb2R5ZWFy
# LmNvbS9Hb29keWVhciUyMFByb2R1Y3Rpb24lMjBSb290JTIwQ0ElMjAyLmNydDAN
# BgkqhkiG9w0BAQsFAAOCAQEAY2W2zzEMmXbEBshTEo6TatOTcelt+o6IOEUBgKIp
# zJj7aG4g/wRNQTMMplbTQdo0BtsQW/wF0A9B3DzoAyodZ1JYFHcil4/icky7ukjs
# sYDog7v4uClRCQc+KdGtMI6hXR7/iVQeFEp0x10PHB9xI5HRJ2kaR7h7uLJrnS4Z
# EVeF9l+xpKtM3m6ep81Hex70sy7+iyPX3CpfgYQ0ndqU7C7emOnWrSV5Yl4dx23W
# v9BnAragjrQ43AD/QDiERGH/eEqI0LL9ECCP8uGPPl1X+5U/WvuYpWc/8XtyZ2Lb
# Ggj8k3pw6XDf8GxH7SJK6njOykz7bhM/GEIlQP7BzX8kqTGCAiswggInAgEBMHkw
# YjEtMCsGA1UEChMkVGhlIEdvb2R5ZWFyIFRpcmUgYW5kIFJ1YmJlciBDb21wYW55
# MTEwLwYDVQQDEyhHb29keWVhciBQcm9kdWN0aW9uIEdlbmVyYWwgUHVycG9zZSBD
# QSAyAhMYAAAWfb+Kc6CcOqMqAAAAABZ9MA0GCWCGSAFlAwQCAQUAoIGEMBgGCisG
# AQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIE0V
# m+jIJqoqVpoZqiYi99GpQuuzLa3Yvx9pfNYjOrR+MA0GCSqGSIb3DQEBAQUABIIB
# AH6KwHEWs1QhhWRdAtt05XiBh+yQyzuDOMKhXZvSK+X8jZaIWlIgKT06Y4HBPWwJ
# T/qRPxdtAcwLm3e2J8CRVpR+0Ae7yccalW3451M9SGNx2+8/O0XIrGAsmdcpNGNW
# syjycNQ+DECO1+/z3asN+WOHJKCX5YgZITyvJmhzg+wCTU62R0CiK/bVDcdC1777
# j3qMluEzwpQI7/k4gC5phibin2Af4UAfckNjHZXUegyxl3KgoCAqyoetC6f9cAHV
# R7n/vr9XAkWBKJ4V6RP91enIEQq1k7kpc2NJttC43wcXkq9BIa46D0TSlqZ4QMN8
# Klikv7SobEkVTn09cvkWRwo=
# SIG # End signature block
