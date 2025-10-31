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
        Write-Output "Using prerelease version for update check..."
        $LatestRelease = (Invoke-RestMethod -Uri ($releaseApiUrl -replace '/latest$', '') -Method Get) | Select-Object -First 1
    }
    else {
        Write-Output "Using stable version for update check..."
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
    # Read the path of the extracted folder
    $extractedSubFolder = Get-Content -Path $updateSignalFile
	
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
    $global:config = Get-Content $global:configFile | ConvertFrom-Json

    # Check if userconfig.json exists
    if (Test-Path -Path $global:userconfigFile -PathType Leaf) {
        # Load user configuration from userconfig.json
        $global:userConfig = Get-Content $global:userConfigFile | ConvertFrom-Json

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
        $allLanguages = Get-Content $LanguageFile | ConvertFrom-Json
        
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

            $settings = Get-Content $settingsFile | ConvertFrom-Json

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
# MIIP2AYJKoZIhvcNAQcCoIIPyTCCD8UCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+F4RljxkFx4w/wg2kFxL5TGq
# 35yggg0oMIIFUzCCBDugAwIBAgITGAAAFn2/inOgnDqjKgAAAAAWfTANBgkqhkiG
# 9w0BAQsFADBiMS0wKwYDVQQKEyRUaGUgR29vZHllYXIgVGlyZSBhbmQgUnViYmVy
# IENvbXBhbnkxMTAvBgNVBAMTKEdvb2R5ZWFyIFByb2R1Y3Rpb24gR2VuZXJhbCBQ
# dXJwb3NlIENBIDIwHhcNMjQwMjI4MTEzNDI5WhcNMjYwMjI3MTEzNDI5WjCBxDEL
# MAkGA1UEBhMCVVMxCzAJBgNVBAgTAk9IMQ4wDAYDVQQHEwVBa3JvbjErMCkGA1UE
# CgwiVGhlIEdvb2R5ZWFyIFRpcmUgJiBSdWJiZXIgQ29tcGFueTE1MDMGA1UECwws
# SG9zdGVkIGJ5IFRoZSBHb29keWVhciBUaXJlICYgUnViYmVyIENvbXBhbnkxDjAM
# BgNVBAsTBUdSSVBTMSQwIgYDVQQDDBtHUklQU19Db2Rlc2lnbi5nb29keWVhci5j
# b20wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDdtqmGEV2HYEO2lMi3
# 3DiYOA/eAQsWdqOP2XjQGFuV2lTGsid4hP3j6q7ZjeKpS1tzWRsSABkZgMgdRW49
# tez9YNf+lK5NOfATr0zspU0b7y06kZS0/rLyTSGVlhFV7rQc3ccizSv4wYiTQP4x
# LAJ93qGHMKJDsN/wFuUOrcqa1LSSXZCxzbJBOzus+YNCWe4ySsnDyWte4xHtGUMI
# oJLZ1uNbk5xot9tPEZV3WGXo1WHQkPLFvuEPMn9BxhFb3NBNa8nE4YZcUjdmt7p6
# KjNwUzUkqLSHlEXj1qNSntXipl/cdiy2vKWx4HtV4kkYoIg1Chtc2HqgewmX/buR
# YhbNAgMBAAGjggGdMIIBmTALBgNVHQ8EBAMCB4AwHQYDVR0OBBYEFEV4A++pnmLy
# 6kKK2YR0HXuOnf4OMB8GA1UdIwQYMBaAFPHdvgAfi3hnJNuB9soFxVlVdAMZMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly9wa2kuZ29vZHllYXIuY29tL0dvb2R5ZWFy
# JTIwUHJvZHVjdGlvbiUyMEdlbmVyYWwlMjBQdXJwb3NlJTIwQ0ElMjAyLmNybDBq
# BggrBgEFBQcBAQReMFwwWgYIKwYBBQUHMAKGTmh0dHA6Ly9wa2kuZ29vZHllYXIu
# Y29tL0dvb2R5ZWFyJTIwUHJvZHVjdGlvbiUyMEdlbmVyYWwlMjBQdXJwb3NlJTIw
# Q0ElMjAyLmNydDAMBgNVHRMBAf8EAjAAMD0GCSsGAQQBgjcVBwQwMC4GJisGAQQB
# gjcVCIXSvkGElMBZhemBJoT07FCGx7V2AITTiB6C/LBMAgFkAgENMBMGA1UdJQQM
# MAoGCCsGAQUFBwMDMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYBBQUHAwMwDQYJKoZI
# hvcNAQELBQADggEBAAGhpGXBNC2vv622nyTK22yNc8WMqBILKQSkd4pIxAont9BL
# L7GrEKsoodwjDbSe8MBys7N3bSFoOP8s+CNxEKBH5hC8FHFNmkF0vCRbx1S8Ki/n
# koE7iZKJg5536KaMXo92dvA5rH7PxnLfUmJnOYEx7Gektj1o/QOq0eNrjADvR4bS
# X9sSqDJkncrg0BBn3k0Vsmz1VDO5IwUF0we2guNXETUmGuWotHb+n0XGxHv5JAfR
# IXjuiz/JOo4EevaGmGuUwNR6Uzg8kWzSN+LXXSnWBLY1t5pRsYEHHXmvCKpVZf7k
# Rj3962/lPWb5khZR0qwzdIL6ZDV66oKLRF5hzeYwggfNMIIGtaADAgECAhMbAAAA
# A50ThRfr1Yg8AAAAAAADMA0GCSqGSIb3DQEBCwUAMFcxLTArBgNVBAoTJFRoZSBH
# b29keWVhciBUaXJlIGFuZCBSdWJiZXIgQ29tcGFueTEmMCQGA1UEAxMdR29vZHll
# YXIgUHJvZHVjdGlvbiBSb290IENBIDIwHhcNMTcwMzA4MTY0MDMxWhcNMzAwMzA4
# MTYyNDUxWjBiMS0wKwYDVQQKEyRUaGUgR29vZHllYXIgVGlyZSBhbmQgUnViYmVy
# IENvbXBhbnkxMTAvBgNVBAMTKEdvb2R5ZWFyIFByb2R1Y3Rpb24gR2VuZXJhbCBQ
# dXJwb3NlIENBIDIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCreYzF
# QzxMuOSutDjyqdC2Q/b5790Ie8bbXAN52NEs3NR3R+NMgbU7b6czOZVHxLPZ4fQj
# Af6aoec6KxpOS7be7PpRsVqp/KfAu2ARO7drpxFu4KZLCitNPzou96xTmJolLrKE
# 5pq4tB9Q+pM+sOC7LSGzxRUzKX463BQDx3qR8R+KccsGEO47+sKGw4E7RRhpJ/gp
# TmU0S+ig1AG/79duMEGVQVR++sSdlBXhPTpboF7/xC6UgkQr6Gs+qjcOIlfkUuRv
# lLsuz4LEz0nsHzdk7HF3sUve1TZNb9t6/8Aykaf58B9Z/3dLYeWTbNyyGY5aFYvD
# LRKz3ZdVYVeNVNaLAgMBAAGjggSFMIIEgTAOBgNVHQ8BAf8EBAMCAQYwEAYJKwYB
# BAGCNxUBBAMCAQAwHQYDVR0OBBYEFPHdvgAfi3hnJNuB9soFxVlVdAMZMIIDOQYD
# VR0gBIIDMDCCAywwggMoBgwrBgEEAYKYEgEBAQEwggMWMIIDEgYIKwYBBQUHAgIw
# ggMEHoIDAABUAGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABpAG8AbgAgAEEA
# dQB0AGgAbwByAGkAdAB5ACAAKABDAEEAKQAgAGkAcwAgAGEAIABHAG8AbwBkAHkA
# ZQBhAHIAIABUAGkAcgBlACAAJgAgAFIAdQBiAGIAZQByACAAQwBvAG0AcABhAG4A
# eQAgAGkAbgB0AGUAcgBuAGEAbAAgAHIAZQBzAG8AdQByAGMAZQAuACAAIABDAGUA
# cgB0AGkAZgBpAGMAYQB0AGUAcwAgAGkAcwBzAHUAZQBkACAAYgB5ACAAdABoAGkA
# cwAgAEMAQQAgAGEAcgBlACAAZgBvAHIAIABpAG4AdABlAHIAbgBhAGwAIABHAG8A
# bwBkAHkAZQBhAHIAIABUAGkAcgBlACAAJgAgAFIAdQBiAGIAZQByACAAQwBvAG0A
# cABhAG4AeQAgAHUAcwBlACAAbwBuAGwAeQAuACAAIABBAG4AeQAgAG4AbwBuAC0A
# RwBvAG8AZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMA
# bwBtAHAAYQBuAHkAIABwAGEAcgB0AHkAIABzAGgAYQBsAGwAIABuAG8AdAAgAHIA
# ZQBsAHkAIABvAG4AIAB0AGgAaQBzACAAQwBBACAAZgBvAHIAIABhAG4AeQAgAHIA
# ZQBhAHMAbwBuAC4AIAAgAEYAbwByACAAbQBvAHIAZQAgAGkAbgBmAG8AcgBtAGEA
# dABpAG8AbgAgAGEAYgBvAHUAdAAgAFQAaABlACAARwBvAG8AZAB5AGUAYQByACAA
# VABpAHIAZQAgACYAIABSAHUAYgBiAGUAcgAgAEMAbwBtAHAAYQBuAHkAIABDAGUA
# cgB0AGkAZgBpAGMAYQB0AGUAIABQAG8AbABpAGMAeQAsACAAcABsAGUAYQBzAGUA
# IABlAG0AYQBpAGwAIABzAGUAYwB1AHIAaQB0AHkAQABnAG8AbwBkAHkAZQBhAHIA
# LgBjAG8AbTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTASBgNVHRMBAf8ECDAG
# AQH/AgEAMB8GA1UdIwQYMBaAFGlQDA9Zyx9ycFtxA/6frb0t9gGcMFIGA1UdHwRL
# MEkwR6BFoEOGQWh0dHA6Ly9wa2kuZ29vZHllYXIuY29tL0dvb2R5ZWFyJTIwUHJv
# ZHVjdGlvbiUyMFJvb3QlMjBDQSUyMDIuY3JsMF0GCCsGAQUFBwEBBFEwTzBNBggr
# BgEFBQcwAoZBaHR0cDovL3BraS5nb29keWVhci5jb20vR29vZHllYXIlMjBQcm9k
# dWN0aW9uJTIwUm9vdCUyMENBJTIwMi5jcnQwDQYJKoZIhvcNAQELBQADggEBAGNl
# ts8xDJl2xAbIUxKOk2rTk3HpbfqOiDhFAYCiKcyY+2huIP8ETUEzDKZW00HaNAbb
# EFv8BdAPQdw86AMqHWdSWBR3IpeP4nJMu7pI7LGA6IO7+LgpUQkHPinRrTCOoV0e
# /4lUHhRKdMddDxwfcSOR0SdpGke4e7iya50uGRFXhfZfsaSrTN5unqfNR3se9LMu
# /osj19wqX4GENJ3alOwu3pjp1q0leWJeHcdt1r/QZwK2oI60ONwA/0A4hERh/3hK
# iNCy/RAgj/Lhjz5dV/uVP1r7mKVnP/F7cmdi2xoI/JN6cOlw3/BsR+0iSup4zspM
# +24TPxhCJUD+wc1/JKkxggIaMIICFgIBATB5MGIxLTArBgNVBAoTJFRoZSBHb29k
# eWVhciBUaXJlIGFuZCBSdWJiZXIgQ29tcGFueTExMC8GA1UEAxMoR29vZHllYXIg
# UHJvZHVjdGlvbiBHZW5lcmFsIFB1cnBvc2UgQ0EgMgITGAAAFn2/inOgnDqjKgAA
# AAAWfTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUQHF4gfMwrSL7jtYLSBySx46VJyEwDQYJKoZI
# hvcNAQEBBQAEggEAeg1tX69dJI1uH8+Tka/l0VeLGy+7Qg3VWH+PMhqHXIc8tLys
# X6ZINxN6R08s2eZyEUtBx7gFlhCOZ0RqQiKn2yRHgOxSjAYxfajbUbvZnnXZzWac
# Ep46ns11JbMv2IEKuJW7v3l9z6fVhrE0Kv85WQeobuQUfnuqxhmF/XHUEzUM1d/U
# pHZqbZj9u4V331IBCCf2RYbZPDQGP40VvcdiMdGSXeprfq7HHLrqhlyJfd9bPZ5P
# NaySGgew1eY1I339WHA0RN+cFurn5O6yMvfabaSc4ecFB2OhFMKhA7X9C7aOPvh9
# pKSDr8QmC4u+xmoE9un8oU6zProQb4bto2GvRg==
# SIG # End signature block
