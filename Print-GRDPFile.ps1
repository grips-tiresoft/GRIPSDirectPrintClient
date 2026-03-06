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
# MII7sgYJKoZIhvcNAQcCoII7ozCCO58CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDypKCbUCZ8AevS
# q6iZE8cbcbGshICYX4tfdqiHEW1aWKCCI9YwggXMMIIDtKADAgECAhBUmNLR1FsZ
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
# 03u4aUoqlmZpxJTG9F9urJh4iIAGXKKy7aIwggdaMIIFQqADAgECAhMzAAAABzeM
# W6HZW4zUAAAAAAAHMA0GCSqGSIb3DQEBDAUAMGMxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xNDAyBgNVBAMTK01pY3Jvc29mdCBJ
# RCBWZXJpZmllZCBDb2RlIFNpZ25pbmcgUENBIDIwMjEwHhcNMjEwNDEzMTczMTU0
# WhcNMjYwNDEzMTczMTU0WjBaMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSswKQYDVQQDEyJNaWNyb3NvZnQgSUQgVmVyaWZpZWQg
# Q1MgQU9DIENBIDAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAt/fA
# AygHxbo+jxA04hNI8bz+EqbWvSu9dRgAawjCZau1Y54IQal5ArpJWi8cIj0WA+mp
# wix8iTRguq9JELZvTMo2Z1U6AtE1Tn3mvq3mywZ9SexVd+rPOTr+uda6GVgwLA80
# LhRf82AvrSwxmZpCH/laT08dn7+Gt0cXYVNKJORm1hSrAjjDQiZ1Jiq/SqiDoHN6
# PGmT5hXKs22E79MeFWYB4y0UlNqW0Z2LPNua8k0rbERdiNS+nTP/xsESZUnrbmyX
# ZaHvcyEKYK85WBz3Sr6Et8Vlbdid/pjBpcHI+HytoaUAGE6rSWqmh7/aEZeDDUkz
# 9uMKOGasIgYnenUk5E0b2U//bQqDv3qdhj9UJYWADNYC/3i3ixcW1VELaU+wTqXT
# xLAFelCi/lRHSjaWipDeE/TbBb0zTCiLnc9nmOjZPKlutMNho91wxo4itcJoIk2b
# Pot9t+AV+UwNaDRIbcEaQaBycl9pcYwWmf0bJ4IFn/CmYMVG1ekCBxByyRNkFkHm
# uMXLX6PMXcveE46jMr9syC3M8JHRddR4zVjd/FxBnS5HOro3pg6StuEPshrp7I/K
# k1cTG8yOWl8aqf6OJeAVyG4lyJ9V+ZxClYmaU5yvtKYKk1FLBnEBfDWw+UAzQV0v
# cLp6AVx2Fc8n0vpoyudr3SwZmckJuz7R+S79BzMCAwEAAaOCAg4wggIKMA4GA1Ud
# DwEB/wQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU6IPEM9fcnwyc
# dpoKptTfh6ZeWO4wVAYDVR0gBE0wSzBJBgRVHSAAMEEwPwYIKwYBBQUHAgEWM2h0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0
# bTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB8GA1UdIwQYMBaAFNlBKbAPD2Ns72nX9c0pnqRIajDmMHAGA1UdHwRpMGcwZaBj
# oGGGX2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29m
# dCUyMElEJTIwVmVyaWZpZWQlMjBDb2RlJTIwU2lnbmluZyUyMFBDQSUyMDIwMjEu
# Y3JsMIGuBggrBgEFBQcBAQSBoTCBnjBtBggrBgEFBQcwAoZhaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBJRCUyMFZlcmlm
# aWVkJTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDIxLmNydDAtBggrBgEFBQcw
# AYYhaHR0cDovL29uZW9jc3AubWljcm9zb2Z0LmNvbS9vY3NwMA0GCSqGSIb3DQEB
# DAUAA4ICAQB3/utLItkwLTp4Nfh99vrbpSsL8NwPIj2+TBnZGL3C8etTGYs+HZUx
# NG+rNeZa+Rzu9oEcAZJDiGjEWytzMavD6Bih3nEWFsIW4aGh4gB4n/pRPeeVrK4i
# 1LG7jJ3kPLRhNOHZiLUQtmrF4V6IxtUFjvBnijaZ9oIxsSSQP8iHMjP92pjQrHBF
# WHGDbkmx+yO6Ian3QN3YmbdfewzSvnQmKbkiTibJgcJ1L0TZ7BwmsDvm+0XRsPOf
# FgnzhLVqZdEyWww10bflOeBKqkb3SaCNQTz8nshaUZhrxVU5qNgYjaaDQQm+P2SE
# pBF7RolEC3lllfuL4AOGCtoNdPOWrx9vBZTXAVdTE2r0IDk8+5y1kLGTLKzmNFn6
# kVCc5BddM7xoDWQ4aUoCRXcsBeRhsclk7kVXP+zJGPOXwjUJbnz2Kt9iF/8B6FDO
# 4blGuGrogMpyXkuwCC2Z4XcfyMjPDhqZYAPGGTUINMtFbau5RtGG1DOWE9edCaht
# uPMDgByfPixvhy3sn7zUHgIC/YsOTMxVuMQi/bgamemo/VNKZrsZaS0nzmOxKpg9
# qDefj5fJ9gIHXcp2F0OHcVwe3KnEXa8kqzMDfrRl/wwKrNSFn3p7g0b44Ad1ONDm
# Wt61MLQvF54LG62i6ffhTCeoFT9Z9pbUo2gxlyTFg7Bm0fgOlnRfGDCCB38wggVn
# oAMCAQICEzMACDILeYiQzbZS0dUAAAAIMgswDQYJKoZIhvcNAQEMBQAwWjELMAkG
# A1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjErMCkGA1UE
# AxMiTWljcm9zb2Z0IElEIFZlcmlmaWVkIENTIEFPQyBDQSAwMTAeFw0yNjAzMDYx
# MjA3NDZaFw0yNjAzMDkxMjA3NDZaMIH8MRMwEQYDVQQREwo0NDMxNi0wMDAxMQsw
# CQYDVQQGEwJVUzENMAsGA1UECBMET2hpbzEOMAwGA1UEBxMFQWtyb24xGzAZBgNV
# BAkTEjIwMCBJbm5vdmF0aW9uIFdheTFNMEsGA1UECh5EAFQAaABlACAARwBvAG8A
# ZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBtAHAA
# YQBuAHkxTTBLBgNVBAMeRABUAGgAZQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQAaQBy
# AGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5MIIBojANBgkqhkiG
# 9w0BAQEFAAOCAY8AMIIBigKCAYEAxvI3+mPNN5fAkfQS/CnvaE3SkaLG4UVqy4v7
# sxtmQioYWG4rvzzU0gsOD+mD2yy9QKl03K0UHrQPimgoS9QVb/hjvPKCmV+riXiv
# OJuiif8PsKbK54etqGt0zZhwTIzInoDWEC6KHxYJ3MnFEOmVkbqxUfonFxU88EcJ
# TCCuuLYpCHaaYlRCWcPQvw+1R25vOcYMkKIkMqPhX/9qnSNNPlxddrrghRbHE6Iv
# +QeRgb8oWqxcCCLDA/J2xtBNAIWrRa+RCzgnP1gIJa5iTMjgDuT1ghpxUg5u6hF2
# rKcexM7SQhdB4CMPVVC1YtW5m4Uh4nIVcL6/SWLgO+pPQEkRp8Vobn6v/WWo8l4P
# Y+T3S5wEi7gM4ZNWPjUK9ch+TjpJWGLuPYffexCcMAXwWgGo33YOm4rf3DXvTw15
# 4XLOWqEhxdhddDr1D3+V4GZJX3S3vFcLtgzisT8VlQlUC8EAPBid6RIyw0ON15WH
# pXvzxTGdP+0BRas5cDCBQ7craZynAgMBAAGjggIZMIICFTAMBgNVHRMBAf8EAjAA
# MA4GA1UdDwEB/wQEAwIHgDA8BgNVHSUENTAzBgorBgEEAYI3YQEABggrBgEFBQcD
# AwYbKwYBBAGCN2GDi6zmXoHCg+Bng6e09AGZrMsTMB0GA1UdDgQWBBSDiwq1kG0k
# NNA5XNa6w2rTGlyLxjAfBgNVHSMEGDAWgBTog8Qz19yfDJx2mgqm1N+Hpl5Y7jBn
# BgNVHR8EYDBeMFygWqBYhlZodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBBT0MlMjBDQSUy
# MDAxLmNybDCBpQYIKwYBBQUHAQEEgZgwgZUwZAYIKwYBBQUHMAKGWGh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSUQlMjBW
# ZXJpZmllZCUyMENTJTIwQU9DJTIwQ0ElMjAwMS5jcnQwLQYIKwYBBQUHMAGGIWh0
# dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDBmBgNVHSAEXzBdMFEGDCsG
# AQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQBMA0GCSqGSIb3
# DQEBDAUAA4ICAQCmSQBcRFUQgwxSSOyCgJNJN8l0KHfhQnM2Tzpr1QTInAE3dqf6
# IaqCNpUyYh3LaoE4ttcYPMlRzgVwwCAX8HUzh+vPgTBSH66Osxfv4ovkwMpeO7GX
# gDWKSTkUk7P5uUvTXmyvM101Mi1WUK9EbcvJRZnP1BuZqFIYAT/Qz+aOKJHOVb9u
# 2VAaTpB+oYzQ++8cTeNZOSxEleT+lMqNztnYpQDDCPSaRHK01pc63ROu/s6AXslY
# bW/pE4keLMXGsOXQazR3rTU5yYG6O61uLmXUzYWouKcjhYg4U7Qf2YggKT5wh9By
# vIeI/3RLmK/gORzqQA3gXgN3doiY4kTs1Y248NtcKfTktYaruhGiH7joErsNT9lZ
# OwaNylS5JR7GEbd7o5y5tGMnL9JW7Pnr7Q6PtJSoTT+Ly77GPYgP1Ndqn2DJnUbz
# UcAAKpVYxXwwv+OcNgSbHmZ1fHiZp22f1TvoWfvZLTXykCFgePLwFogrRqTMupgU
# gIEP5uBHu6TfGFelKK8Hg35V/7PfZ7sXvPePLqGCu0TsJf/yHd1Hp0ODbDSooqdO
# NOXm2bVBlI7RPeRpjFSKg0uQIRZON27NdQdukyOB+gwncNDLWyktgtlMe43zICVq
# pWAQpy9ZKzYVh4PE6lA2y9pHz1YzHlG2W5tx1RKBVeNIrwZl9m53npiBUjCCB38w
# ggVnoAMCAQICEzMACDILeYiQzbZS0dUAAAAIMgswDQYJKoZIhvcNAQEMBQAwWjEL
# MAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjErMCkG
# A1UEAxMiTWljcm9zb2Z0IElEIFZlcmlmaWVkIENTIEFPQyBDQSAwMTAeFw0yNjAz
# MDYxMjA3NDZaFw0yNjAzMDkxMjA3NDZaMIH8MRMwEQYDVQQREwo0NDMxNi0wMDAx
# MQswCQYDVQQGEwJVUzENMAsGA1UECBMET2hpbzEOMAwGA1UEBxMFQWtyb24xGzAZ
# BgNVBAkTEjIwMCBJbm5vdmF0aW9uIFdheTFNMEsGA1UECh5EAFQAaABlACAARwBv
# AG8AZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBt
# AHAAYQBuAHkxTTBLBgNVBAMeRABUAGgAZQAgAEcAbwBvAGQAeQBlAGEAcgAgAFQA
# aQByAGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEAbgB5MIIBojANBgkq
# hkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAxvI3+mPNN5fAkfQS/CnvaE3SkaLG4UVq
# y4v7sxtmQioYWG4rvzzU0gsOD+mD2yy9QKl03K0UHrQPimgoS9QVb/hjvPKCmV+r
# iXivOJuiif8PsKbK54etqGt0zZhwTIzInoDWEC6KHxYJ3MnFEOmVkbqxUfonFxU8
# 8EcJTCCuuLYpCHaaYlRCWcPQvw+1R25vOcYMkKIkMqPhX/9qnSNNPlxddrrghRbH
# E6Iv+QeRgb8oWqxcCCLDA/J2xtBNAIWrRa+RCzgnP1gIJa5iTMjgDuT1ghpxUg5u
# 6hF2rKcexM7SQhdB4CMPVVC1YtW5m4Uh4nIVcL6/SWLgO+pPQEkRp8Vobn6v/WWo
# 8l4PY+T3S5wEi7gM4ZNWPjUK9ch+TjpJWGLuPYffexCcMAXwWgGo33YOm4rf3DXv
# Tw154XLOWqEhxdhddDr1D3+V4GZJX3S3vFcLtgzisT8VlQlUC8EAPBid6RIyw0ON
# 15WHpXvzxTGdP+0BRas5cDCBQ7craZynAgMBAAGjggIZMIICFTAMBgNVHRMBAf8E
# AjAAMA4GA1UdDwEB/wQEAwIHgDA8BgNVHSUENTAzBgorBgEEAYI3YQEABggrBgEF
# BQcDAwYbKwYBBAGCN2GDi6zmXoHCg+Bng6e09AGZrMsTMB0GA1UdDgQWBBSDiwq1
# kG0kNNA5XNa6w2rTGlyLxjAfBgNVHSMEGDAWgBTog8Qz19yfDJx2mgqm1N+Hpl5Y
# 7jBnBgNVHR8EYDBeMFygWqBYhlZodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBBT0MlMjBD
# QSUyMDAxLmNybDCBpQYIKwYBBQUHAQEEgZgwgZUwZAYIKwYBBQUHMAKGWGh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSUQl
# MjBWZXJpZmllZCUyMENTJTIwQU9DJTIwQ0ElMjAwMS5jcnQwLQYIKwYBBQUHMAGG
# IWh0dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDBmBgNVHSAEXzBdMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQBMA0GCSqG
# SIb3DQEBDAUAA4ICAQCmSQBcRFUQgwxSSOyCgJNJN8l0KHfhQnM2Tzpr1QTInAE3
# dqf6IaqCNpUyYh3LaoE4ttcYPMlRzgVwwCAX8HUzh+vPgTBSH66Osxfv4ovkwMpe
# O7GXgDWKSTkUk7P5uUvTXmyvM101Mi1WUK9EbcvJRZnP1BuZqFIYAT/Qz+aOKJHO
# Vb9u2VAaTpB+oYzQ++8cTeNZOSxEleT+lMqNztnYpQDDCPSaRHK01pc63ROu/s6A
# XslYbW/pE4keLMXGsOXQazR3rTU5yYG6O61uLmXUzYWouKcjhYg4U7Qf2YggKT5w
# h9ByvIeI/3RLmK/gORzqQA3gXgN3doiY4kTs1Y248NtcKfTktYaruhGiH7joErsN
# T9lZOwaNylS5JR7GEbd7o5y5tGMnL9JW7Pnr7Q6PtJSoTT+Ly77GPYgP1Ndqn2DJ
# nUbzUcAAKpVYxXwwv+OcNgSbHmZ1fHiZp22f1TvoWfvZLTXykCFgePLwFogrRqTM
# upgUgIEP5uBHu6TfGFelKK8Hg35V/7PfZ7sXvPePLqGCu0TsJf/yHd1Hp0ODbDSo
# oqdONOXm2bVBlI7RPeRpjFSKg0uQIRZON27NdQdukyOB+gwncNDLWyktgtlMe43z
# ICVqpWAQpy9ZKzYVh4PE6lA2y9pHz1YzHlG2W5tx1RKBVeNIrwZl9m53npiBUjCC
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
# BgNVBAMTIk1pY3Jvc29mdCBJRCBWZXJpZmllZCBDUyBBT0MgQ0EgMDECEzMACDIL
# eYiQzbZS0dUAAAAIMgswDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3AgEMMQIw
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQgTRWb
# 6MgmqipWmhmqJiL30alC67Mtrdi/H2l81iM6tH4wDQYJKoZIhvcNAQEBBQAEggGA
# sCzRgWL9AnNKgI2yqg1IftZ0nOHg/f70KVyOC+NK/EFDWgdZ8+Ej+iQsqOzq3LxF
# bCjbqd0KjSuAdNeTuU6Xh8JQo6SXYTn1roHIeDcUz6ysPJUQk6QqgVFughu/ZDli
# ZLXHHVSSlKtx7z781uO5knCESb4We6A/39Z6oCxPiAo2JUYpn91ceCcQOqrU9DAY
# ABZFscfeSe0TYt+/LdKBlrrg7ZEM7Poi+9i+9bnNc4L87eR/Q/9fW/nBRARoDDp6
# tto5JW/GItSH0nTyXRK8rL2d40LdaYyHQXeaMlYP7PMe0wUMSXEL1FP/fUWRjB+f
# EetSkuUBHZ9GX0sDnIPseQvslNnHlhd7gev5LpOfZfGqACc+tbl2dbGv7A9MPonX
# be2KCrz0I8EN7yfO4rAp5zv41AVu1G8hzyS/7EnyCGqeOCqGivaICBVqeiov448h
# 1mBzqZg7keTEz6IEn+KDsW2xPpidnSLmjnLtIKIbduaP575UeNOQvZHqLr2znHjZ
# oYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIUizCCFIcC
# AQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSCAVUwggFR
# AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIFQndGKnwX5Du4bX718Z
# 7pJkjv7TlcKRXZ1aDA3Qg4OjAgZpoXHn5SMYEzIwMjYwMzA2MTQzNTMxLjU3MVow
# BIACAfSggemkgeYwgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRl
# ZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjQ1MUEtMDVFMC1EOTQ3MTUwMwYD
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
# JYbYbBh7pmgAXVswggefMIIFh6ADAgECAhMzAAAAXGFMAKan8ukjAAAAAABcMA0G
# CSqGSIb3DQEBDAUAMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMgUlNBIFRpbWVz
# dGFtcGluZyBDQSAyMDIwMB4XDTI2MDEwODE4NTkwNloXDTI3MDEwNzE4NTkwNlow
# geMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsT
# JE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMe
# blNoaWVsZCBUU1MgRVNOOjQ1MUEtMDVFMC1EOTQ3MTUwMwYDVQQDEyxNaWNyb3Nv
# ZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1dGhvcml0eTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBAK41jM5VeZNHd4NtVlUxNoOHqf6e/TgnDjNb
# q/s5M3WhNLEEiHKrpAowFqcDZ5ek7KpDRW9xeBwLlLPdoyCV1fIpHv1swh4tU9C1
# dIKyOusuN32al13lVCrgd8qBqKw8sW/rKenXHaDW94g2HlX5ssQWnVcKpZYIR3Up
# ExCmRSQtIVEFloWZvo3PcwatECZTifG3TAhX/bTBKXOFoUpqWk/SFLAj7WxULE3w
# fyuWCH7SMY1mAIkSdcrDdI7/AkSNGFkw9pRIYRftzAKvOqP/LWGcD5rwZvB9nSUJ
# A+YyuNty/HePVUvpV5wWel2F/I/kCksfkJKRe01iGIWLuYbzz34yLP4JyvtrA+IU
# /Uu8BH0JMNkfPi5r8DmviM56Ef0BTt4Q/G6UDDLQEHLA6Xu2kreAox7HBDKAcWll
# 7l3tM+K5XMrCS1uH2ujgKQPByl1QcJGDKw065mvxZ3rR/0MuSKihAVGH+aJxnHzY
# NRMOUzkK1yXg2zgd6Nk1Eav4a93tX2Qq5kXpeZba132ohkYvh5vnXkjQVQ9HJwfe
# +J+09lUpgNuSHwb/gjzi2Z8+Hg9pwe7okJ2QLQGw5TQRkw/xK7aH02aDBXOTzo1U
# ft/UmEVxbCzFnGFGAjcr0RwfnGFaIyyrF5uZmSsWKCF2o7v7LOjNlneWSSoFaKRh
# ODuKRjw/AgMBAAGjggHLMIIBxzAdBgNVHQ4EFgQUrvsyHJEYZEQ5bnTgWgMSJ7ok
# 52kwHwYDVR0jBBgwFoAUa2koOjUvSGNAz3vYr0npPtk92yEwbAYDVR0fBGUwYzBh
# oF+gXYZbaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9z
# b2Z0JTIwUHVibGljJTIwUlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNy
# bDB5BggrBgEFBQcBAQRtMGswaQYIKwYBBQUHMAKGXWh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwUHVibGljJTIwUlNBJTIw
# VGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNydDAMBgNVHRMBAf8EAjAAMBYGA1Ud
# JQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDBmBgNVHSAEXzBdMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQCMA0GCSqG
# SIb3DQEBDAUAA4ICAQCFdtA33LEkhAWdgIW55V5nT1ZG/VaQ4M2yBBDuIKUf20fP
# W0PqluLQtPcodrXGZDeCiON3Op3kUTXAYwBpm+7YPMrJbRlZgMQPsFtbSdHC9bET
# 0NJV8oXDGEWBOQe7lwwH05+Pa17xB2/mmVenWWtrZX11I8BoFjcP2DwZSJ5X1G6T
# Te6GYJl99H58ihs7Tv/TEZP7c8mj6eVeOCsiyKGXpRTYJL9yltt6cKKccmcraVSj
# MTv43bmpLLL1WCq+y/T3HEtjXizKXRo5m3TE69UgiosqZ66LeH0dvFB9u1+B7EtD
# jTPvCyc4n83+W3vm9BqzBmvf3ESgwt4J5KHK+CoyKieaRBdP3rVOLouk3MaFkAFH
# thWlf6rzh/0DopuXxqYby3Eot1c530MqQJ31v72xJMSe5Qs+zK55qsKg22vck6b2
# mwhD/ZNh66YeBsVY2kS9ahkGMtR+OgPz0DnDi+OeR60bwiZHJifwyHMTcGuT9pUn
# p7nCt6e0vIhhtmJ1tlEIMurW4SYtUKiULzP6SLdq8ys1+BUohCht/tgi+ibWfTSn
# o7veBC7xaJi6JskqA/E8PdEzUOWHabwYiSHXlZMbobOTR0Dk3IzaqgUUX+E2xsfm
# Vj2outhvWZ55c1/PfLK1p4yzRe9+rY3V4Ni0ORGUQWhefgkNYspB1jZiXDq8gzGC
# A9QwggPQAgEBMHgwYTELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFB1YmxpYyBSU0EgVGltZXN0
# YW1waW5nIENBIDIwMjACEzMAAABcYUwApqfy6SMAAAAAAFwwDQYJYIZIAWUDBAIB
# BQCgggEtMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQg485sZcA88280laPUlYrqgkU9bgPWpc8hWyG3xLegnCAwgd0GCyqGSIb3DQEJ
# EAIvMYHNMIHKMIHHMIGgBCD3umpBh3duMPkt6ItkJtOzRwBRt1Q8I513q8lrfqDu
# 3DB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lc3RhbXBp
# bmcgQ0EgMjAyMAITMwAAAFxhTACmp/LpIwAAAAAAXDAiBCA5dfFHtIbRwbOK3MDJ
# gnNXdO9/T/9KCAMvLke1NpBe6jANBgkqhkiG9w0BAQsFAASCAgBapfgR4oBCOjT+
# NTfaAvCi5Udwvlt/D329XoDbLv1Cs+8aRBXnlwWhxbzKQ351zuWfUegbyJvbGD/r
# 9TZvBMVYBBMMzkcMrX/IqgSrvhZaMKU6yc+5cWxmhqnTgqqdiNF+fsXKZE0SgI8a
# cfJ6GiR05DDLeoamYf3MCeytBd7kSuDX6PcT6ChEK0EkMQ/SOM6pU7qM+HvMc0eF
# Dh/pBK7jecp4fGk3ZyzCNcmdM6q43xlSR1t6eNM6EHDN7zGGfHPuloD3rNe2xooG
# MQR7UQNHmJhL7U/f3yCoQZLpTSgZrlUARBbDU5Unv9gkRcYPO54r9WLIjGdizwp3
# zOLHSuai+krBzikd0koFriT7GrS245a7/fWmtdBkiczvqOb6+sh9Ew+aiMfyQqMe
# +1jaCvd1SD2qivu9vC2RT3msTPgAGWkDhjnKkR6576gmT61K7Ix5FUG70mJUa9xT
# vrwFvzVbO4E6PVSjfBvRw5K2Nw4HgmgohABYn8NdnP8Tc/VPwETuehr9bbgCO21J
# CRb8pmgrHW+gX3mqc3TOkl7WzUTeABAGqv1EXY511ldbEKS5+BSyHmHPc6wX4ccS
# jCZ1+Dba4AYC4EjA1Rss8uKbxMIxlGIBEoNLb2yOdWidZxsd92F774pPLKMRZc47
# jZKliPPoxIvQ1pExMTMpfNYktFme6Q==
# SIG # End signature block
