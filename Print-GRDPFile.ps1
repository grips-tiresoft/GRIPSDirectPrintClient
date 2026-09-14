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

# Resolve and create the app inbox folder under LocalAppData.
function Get-InboxFolder {
    $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw "Unable to resolve LocalApplicationData folder."
    }

    $inboxFolder = Join-Path -Path $localAppData -ChildPath 'Downloads-GRIPSDirectPrint'
    if (-not (Test-Path -Path $inboxFolder -PathType Container)) {
        New-Item -Path $inboxFolder -ItemType Directory -Force | Out-Null
    }

    return $inboxFolder
}

# Resolve the user's Downloads folder using Known Folder APIs with fallbacks.
function Get-DownloadsFolder {
    try {
        $downloadsEnum = [System.Enum]::Parse([System.Environment+SpecialFolder], 'Downloads')
        $downloadsPath = [Environment]::GetFolderPath($downloadsEnum)
        if (-not [string]::IsNullOrWhiteSpace($downloadsPath) -and (Test-Path -Path $downloadsPath -PathType Container)) {
            return $downloadsPath
        }
    }
    catch {
    }

    try {
        $shell = New-Object -ComObject Shell.Application
        $downloadsFolder = $shell.Namespace('shell:Downloads')
        if ($null -ne $downloadsFolder -and $downloadsFolder.Self -and -not [string]::IsNullOrWhiteSpace($downloadsFolder.Self.Path)) {
            return $downloadsFolder.Self.Path
        }
    }
    catch {
    }

    try {
        $userShellFoldersPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
        $downloadsKnownFolderId = '{374DE290-123F-4565-9164-39C4925E467B}'
        $downloadsPath = (Get-ItemProperty -Path $userShellFoldersPath -Name $downloadsKnownFolderId -ErrorAction Stop).$downloadsKnownFolderId

        if (-not [string]::IsNullOrWhiteSpace($downloadsPath)) {
            $expandedPath = [Environment]::ExpandEnvironmentVariables($downloadsPath)
            if (Test-Path -Path $expandedPath -PathType Container) {
                return $expandedPath
            }
        }
    }
    catch {
    }

    return (Join-Path -Path ([Environment]::GetFolderPath('UserProfile')) -ChildPath 'Downloads')
}

# Function to check if a printer exists
function Test-PrinterExists {
    param([string]$PrinterName)
    try {
        $printer = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $PrinterName }
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
        [PSCustomObject]$LanguageStrings,
        [PSCustomObject]$Config
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
    
    # Check if there's a redirected or mapped printer from RDP/Citrix
    $redirectedPrinter = $null
    if ($Config -and $Config.PrinterRedirectionSuffixes) {
        $suffixPattern = ($Config.PrinterRedirectionSuffixes | ForEach-Object { [regex]::Escape($_) }) -join "|"
        $redirectedPrinter = $printers | Where-Object {
            $_ -match "^$([regex]::Escape($MissingPrinterName))\s+\((?:$suffixPattern)\s+"
        } | Select-Object -First 1
    }
    
    if ($redirectedPrinter) {
        return $redirectedPrinter
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
                        $alternativePrinter = Select-AlternativePrinter -MissingPrinterName $printer -LanguageStrings $global:LanguageStrings -Config $global:config
                        
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
                    $inboxFolder = Get-InboxFolder
                    $uniqueFilePath = Get-UniqueFileName -FilePath (Join-Path -Path $inboxFolder -ChildPath ([System.IO.Path]::GetFileName($filePath)))
                    Copy-Item -Path $filePath -Destination $uniqueFilePath
                    Write-Host "Opening file: $uniqueFilePath"
                    Start-Process -FilePath $uniqueFilePath
                    continue
                }
            }        
        }
        finally {
            # Clean up temp folder
            Remove-Item -Path $tempFolder -Recurse -Force
            
            # Remove old inbox files
            $inboxFolder = Get-InboxFolder

            if (Test-Path -Path $inboxFolder -PathType Container) {
                # Remove old .eml files
                $inboxPath = Join-Path -Path $inboxFolder -ChildPath "NewEmail*.eml"
                $inboxFiles = Get-ChildItem -Path $inboxPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
                if ($null -ne $inboxFiles) { 
                    Write-Host "Removing old .eml files:" 
                    Write-Host $inboxFiles
                    $inboxFiles | Remove-Item -Force 
                }

                # Remove old .sig files
                $inboxPath = Join-Path -Path $inboxFolder -ChildPath "*.sig"
                $inboxFiles = Get-ChildItem -Path $inboxPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
                if ($null -ne $inboxFiles) {
                    Write-Host "Removing old .sig files:"
                    Write-Host $inboxFiles
                    $inboxFiles | Remove-Item -Force
                }

                # Remove old .grdp files from inbox
                $inboxPath = Join-Path -Path $inboxFolder -ChildPath "*.grdp"
                $inboxFiles = Get-ChildItem -Path $inboxPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
                if ($null -ne $inboxFiles) {
                    Write-Host "Removing old .grdp files:"
                    Write-Host $inboxFiles
                    $inboxFiles | Remove-Item -Force
                }
            }
            else {
                Write-Warning "Inbox folder not found, skipping inbox cleanup: $inboxFolder"
            }

            # Also remove old .grdp files from Downloads to clean browser-origin files.
            $downloadsFolder = Get-DownloadsFolder
            if (Test-Path -Path $downloadsFolder -PathType Container) {
                $downloadsPath = Join-Path -Path $downloadsFolder -ChildPath "*.grdp"
                $downloads = Get-ChildItem -Path $downloadsPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
                if ($null -ne $downloads) {
                    Write-Host "Removing old .grdp files from Downloads:"
                    Write-Host $downloads
                    $downloads | Remove-Item -Force
                }
            }
            else {
                Write-Warning "Downloads folder not found, skipping .grdp cleanup: $downloadsFolder"
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
# MII6+AYJKoZIhvcNAQcCoII66TCCOuUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJELvP2mFTYHA5
# ePILIGsdchJhhKtRNAp9IJieSxQcZ6CCIxwwggXMMIIDtKADAgECAhBUmNLR1FsZ
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
# 03u4aUoqlmZpxJTG9F9urJh4iIAGXKKy7aIwggcoMIIFEKADAgECAhMzAAAAFydF
# CQuLh6/GAAAAAAAXMA0GCSqGSIb3DQEBDAUAMGMxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xNDAyBgNVBAMTK01pY3Jvc29mdCBJ
# RCBWZXJpZmllZCBDb2RlIFNpZ25pbmcgUENBIDIwMjEwHhcNMjYwMzI2MTgxMTMx
# WhcNMzEwMzI2MTgxMTMxWjBaMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSswKQYDVQQDEyJNaWNyb3NvZnQgSUQgVmVyaWZpZWQg
# Q1MgRU9DIENBIDA0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAgsdk
# /gMPZioBlcyfk6tDzJ+PRt4rSLGKW8ewpS0kRxXtURC3T3GdbCKljobEn8ussqhG
# qQpRh/SXvRVwNXEIGb76UG5IPkCJ1S6/9BD61QQsKzPepW0SNj8TXgsFxvS7Mlto
# RuikIIp7Q5jQgaOM6QyK9++6ZVXUpYmZulAe6x8JrwZ0dNkE+rZ66lqtoocwepUS
# VUxM7odDmn8yDHjJ2DNPsfr3uRDix3X4qvh14jH/SW+2Cx7WIMhyIiQO201i6hUi
# xmk4e2ZW8W7C1wPdTjq6BKb+zo8xbrt7ZKQvRX5QOA6dhLquPqj5sVKnxqfk19IC
# 0SafTSTs8yC43Ew965BRRW8VL9ccoOmr4rxQy7aCgYTNk3dd/LphNaTTmnGp7kmL
# TxyHkB5geoWhYuuGrywS8E0wJv0W4rfOtHBV0e9sKvuUIeIUpnsx6ilxEVj6VQXv
# gD6yeCKnPmj3jJiJKAlmUDtth5yzRVBUl44sMiG4L5R/yyACRKk2n088Q2YCoZS1
# O86+oMLKt1jaXGECOjbsVp8Id1VQw8he6J0KirOS5e25XlTdGPFb6oBOOaacgW78
# Kjf0bp+XzAgkc92mDGNJGYSjvdnj+7eMx6meW0DAIGdLRNj8/429MIspFBfz3KDq
# qpN71S4kQ2LLer3dxhDDczKVFL0HLwRuOvgjiG8CAwEAAaOCAdwwggHYMA4GA1Ud
# DwEB/wQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUmvFUd3UMhxY3
# RqCs3nn59H/BeOkwVAYDVR0gBE0wSzBJBgRVHSAAMEEwPwYIKwYBBQUHAgEWM2h0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0
# bTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB8GA1UdIwQYMBaAFNlBKbAPD2Ns72nX9c0pnqRIajDmMHAGA1UdHwRpMGcwZaBj
# oGGGX2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29m
# dCUyMElEJTIwVmVyaWZpZWQlMjBDb2RlJTIwU2lnbmluZyUyMFBDQSUyMDIwMjEu
# Y3JsMH0GCCsGAQUFBwEBBHEwbzBtBggrBgEFBQcwAoZhaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVk
# JTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDIxLmNydDANBgkqhkiG9w0BAQwF
# AAOCAgEAkHVaGf1NJt/JdoimmRZbMWr6baaDi8mkdWvWStk0hdZDpxSYTA7HuipA
# oLL3qIhI101XOl7fOiCh5++jZOamQdAV79ojEUNoIgCZmL2XJrLaGanwdjNynecJ
# yYVCTrRf2+h7KknpWOp4axdOs6K9ZQ5g0IsQWXCwfc0dfkSkLKNY3pDcWLlJPh2j
# d5NUue6pNDv/2G5MFNJhCwltODebyAjGceU+XOzav+7i721YQnQ+39m2aQOFO7zp
# AdaKAeAGhEd6Y6CdDGneSxcoujWvafWbv4ay3jo1ORSLUuWMbKr5X18QE4Sde+gp
# pGLLSkZsrUh2eyYSkX1envWX7ZPzg2/wiuKRlQFarDn+N9+20BqzhxwkNyLzfYJp
# 1Lg4fCXb24XqFjx8SDdRgebFImOfOLVze8XQ/CwkrEaib0PHu2t4GVk4FYroEbNU
# FqvjdBvTY3uiR5TdQoyXoYHvh+TxpLSY2vo7hhK9D/rpEpHC+qmmcRUE4d0gyO9Z
# b1vvt25fxM3ekjvDfVHcPq3qMr0Rwsk4krKZWUEgU1SXT5qN6gqRrshxbT6OQgZ9
# /xT04qiXdzPQR6KindBvSpoOnxnALxcJyzVwNpKL+9u8EZYy98qX6i+4gE/2J6cb
# pekcB0ZXDn/XQxoNUUb6/djT/wllVyG+vIHkdq71PzbH5rYxdcAwggc7MIIFI6AD
# AgECAhMzAAYKKCXSsQajNhxFAAAABgooMA0GCSqGSIb3DQEBDAUAMFoxCzAJBgNV
# BAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKzApBgNVBAMT
# Ik1pY3Jvc29mdCBJRCBWZXJpZmllZCBDUyBFT0MgQ0EgMDQwHhcNMjYwOTEyMjIx
# ODE1WhcNMjYwOTE1MjIxODE1WjCB/DETMBEGA1UEERMKNDQzMTYtMDAwMTELMAkG
# A1UEBhMCVVMxDTALBgNVBAgTBE9oaW8xDjAMBgNVBAcTBUFrcm9uMRswGQYDVQQJ
# ExIyMDAgSW5ub3ZhdGlvbiBXYXkxTTBLBgNVBAoeRABUAGgAZQAgAEcAbwBvAGQA
# eQBlAGEAcgAgAFQAaQByAGUAIAAmACAAUgB1AGIAYgBlAHIAIABDAG8AbQBwAGEA
# bgB5MU0wSwYDVQQDHkQAVABoAGUAIABHAG8AbwBkAHkAZQBhAHIAIABUAGkAcgBl
# ACAAJgAgAFIAdQBiAGIAZQByACAAQwBvAG0AcABhAG4AeTCCAaIwDQYJKoZIhvcN
# AQEBBQADggGPADCCAYoCggGBALqdCWxsRvt9srd1H3WcD1BvGjjpJ3WrF/KWTUBI
# R0ZtkNcRmlfhR+EsPMQX74nQG6MOSsfL+LTS2BpKTKxLQPHczdvn1Mwj2hT2Ttpo
# IH1iFHfvMrQnB5kliJYQBvAK6Kb+z6nlURRgZpjDet1jrDEM2wcEMgB3326lvexF
# Mr7BgESCqMartoJQIKy2xfaB/QaLdaBZrMObMwbqzxPDgDjpE8pszQ+3mEXtGsoh
# 4JHSFAM/4nV8LVbN0+KXNg2oNCjHpfFDThYfLGLPAaXWBBM8w7dSbB1vU6YyBi2u
# 9aSkSuymWwzLC/jukEexRy3Nd4rMgrrCMjfVO2cGnJoXfHDT1LcFAfIDk6Z4sh9a
# /4wfRf99PapcPl0uAbE9fOsaJ5/Yv/a4mJGrfMyKTmW3+UuMQg/x7WLxzA++/AKC
# wId8zQD1qGZdRHgf4BfK3rbHWvL5rqSyZ1e42x5uEaRJw1ZY7r+Mp6RoMvk7m8AE
# 08Jm7mNq/5yfxMaCMD+3yKTRowIDAQABo4IB1TCCAdEwDAYDVR0TAQH/BAIwADAO
# BgNVHQ8BAf8EBAMCB4AwPAYDVR0lBDUwMwYKKwYBBAGCN2EBAAYIKwYBBQUHAwMG
# GysGAQQBgjdhg4us5l6BwoPgZ4OntPQBmazLEzAdBgNVHQ4EFgQUEKz1eCd9yAbU
# zXNU+Rwz1SJ5N10wHwYDVR0jBBgwFoAUmvFUd3UMhxY3RqCs3nn59H/BeOkwZwYD
# VR0fBGAwXjBcoFqgWIZWaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljcm9zb2Z0JTIwSUQlMjBWZXJpZmllZCUyMENTJTIwRU9DJTIwQ0ElMjAw
# NC5jcmwwdAYIKwYBBQUHAQEEaDBmMGQGCCsGAQUFBzAChlhodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMElEJTIwVmVyaWZp
# ZWQlMjBDUyUyMEVPQyUyMENBJTIwMDQuY3J0MFQGA1UdIARNMEswSQYEVR0gADBB
# MD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0Rv
# Y3MvUmVwb3NpdG9yeS5odG0wDQYJKoZIhvcNAQEMBQADggIBAAEs9SrLqgDwJU4G
# Vg/nZpmvyp1LRswI8ngN/E/OOT8pDG51rQaR+M3zv2qmljuGs2LDx95pb5z0m5fS
# 7eU21isnTa4Aa+2iLXw1/8TBA2QSEDjnUJmbk+DQYm+yI6heSamIpk7UMb5+oRI3
# jtj+jL1LpmmZFWTlA23T8D1+MMNo3fEfOrhvbEG44YmPkYmuOaxuFvJkypahErhW
# TvLf83q8Wh6+JA5zV9Chn+HkIDGCX6eUd/8k7iiWdTxSEm7Qev9f4yD0s5ctQzYS
# 6rHMv1eKdy3/5lyhlUigSFihsKVzgCQ0xFrkJrCJh1qQv7rujbMrjo78EEofIeVV
# Ij549ay/J6QvZs8c7teIgARd/YWl1BAaqOYpCtPMd4nNBlxHV2bWSCFwwcfMQiBk
# Uzu4DSqN5+g1JRSUTdtz334REhzaPx5zpdF9738WgDDIghsZPREz2hrZsZIy+wrp
# eIos0DdKbaGqGdbJP/bVNvG8Y669tETn9XUoP6BFnE2l1NygIUc9p0BY2VSnqDol
# ksmFdqJEEQRniilEeJYM9KVDey4bw4/1/aLIaALUSOhU9rc6e3usZUKY3XvPChQM
# jo8GE46VVcCVn6rg0SLfEvzUGBv1jiF6Q/lLTcPh8vTKDMmrrD9oK3wgdsX172Rj
# EVfH/Qc/7GKQAIAkgrTQJiRp6RkrMIIHOzCCBSOgAwIBAgITMwAGCigl0rEGozYc
# RQAAAAYKKDANBgkqhkiG9w0BAQwFADBaMQswCQYDVQQGEwJVUzEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSswKQYDVQQDEyJNaWNyb3NvZnQgSUQgVmVy
# aWZpZWQgQ1MgRU9DIENBIDA0MB4XDTI2MDkxMjIyMTgxNVoXDTI2MDkxNTIyMTgx
# NVowgfwxEzARBgNVBBETCjQ0MzE2LTAwMDExCzAJBgNVBAYTAlVTMQ0wCwYDVQQI
# EwRPaGlvMQ4wDAYDVQQHEwVBa3JvbjEbMBkGA1UECRMSMjAwIElubm92YXRpb24g
# V2F5MU0wSwYDVQQKHkQAVABoAGUAIABHAG8AbwBkAHkAZQBhAHIAIABUAGkAcgBl
# ACAAJgAgAFIAdQBiAGIAZQByACAAQwBvAG0AcABhAG4AeTFNMEsGA1UEAx5EAFQA
# aABlACAARwBvAG8AZAB5AGUAYQByACAAVABpAHIAZQAgACYAIABSAHUAYgBiAGUA
# cgAgAEMAbwBtAHAAYQBuAHkwggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAwggGKAoIB
# gQC6nQlsbEb7fbK3dR91nA9Qbxo46Sd1qxfylk1ASEdGbZDXEZpX4UfhLDzEF++J
# 0BujDkrHy/i00tgaSkysS0Dx3M3b59TMI9oU9k7aaCB9YhR37zK0JweZJYiWEAbw
# Cuim/s+p5VEUYGaYw3rdY6wxDNsHBDIAd99upb3sRTK+wYBEgqjGq7aCUCCstsX2
# gf0Gi3WgWazDmzMG6s8Tw4A46RPKbM0Pt5hF7RrKIeCR0hQDP+J1fC1WzdPilzYN
# qDQox6XxQ04WHyxizwGl1gQTPMO3Umwdb1OmMgYtrvWkpErsplsMywv47pBHsUct
# zXeKzIK6wjI31TtnBpyaF3xw09S3BQHyA5OmeLIfWv+MH0X/fT2qXD5dLgGxPXzr
# Gief2L/2uJiRq3zMik5lt/lLjEIP8e1i8cwPvvwCgsCHfM0A9ahmXUR4H+AXyt62
# x1ry+a6ksmdXuNsebhGkScNWWO6/jKekaDL5O5vABNPCZu5jav+cn8TGgjA/t8ik
# 0aMCAwEAAaOCAdUwggHRMAwGA1UdEwEB/wQCMAAwDgYDVR0PAQH/BAQDAgeAMDwG
# A1UdJQQ1MDMGCisGAQQBgjdhAQAGCCsGAQUFBwMDBhsrBgEEAYI3YYOLrOZegcKD
# 4GeDp7T0AZmsyxMwHQYDVR0OBBYEFBCs9XgnfcgG1M1zVPkcM9UieTddMB8GA1Ud
# IwQYMBaAFJrxVHd1DIcWN0agrN55+fR/wXjpMGcGA1UdHwRgMF4wXKBaoFiGVmh0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMElE
# JTIwVmVyaWZpZWQlMjBDUyUyMEVPQyUyMENBJTIwMDQuY3JsMHQGCCsGAQUFBwEB
# BGgwZjBkBggrBgEFBQcwAoZYaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jZXJ0cy9NaWNyb3NvZnQlMjBJRCUyMFZlcmlmaWVkJTIwQ1MlMjBFT0MlMjBD
# QSUyMDA0LmNydDBUBgNVHSAETTBLMEkGBFUdIAAwQTA/BggrBgEFBQcCARYzaHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnkuaHRt
# MA0GCSqGSIb3DQEBDAUAA4ICAQABLPUqy6oA8CVOBlYP52aZr8qdS0bMCPJ4DfxP
# zjk/KQxuda0GkfjN879qppY7hrNiw8feaW+c9JuX0u3lNtYrJ02uAGvtoi18Nf/E
# wQNkEhA451CZm5Pg0GJvsiOoXkmpiKZO1DG+fqESN47Y/oy9S6ZpmRVk5QNt0/A9
# fjDDaN3xHzq4b2xBuOGJj5GJrjmsbhbyZMqWoRK4Vk7y3/N6vFoeviQOc1fQoZ/h
# 5CAxgl+nlHf/JO4olnU8UhJu0Hr/X+Mg9LOXLUM2EuqxzL9Xinct/+ZcoZVIoEhY
# obClc4AkNMRa5CawiYdakL+67o2zK46O/BBKHyHlVSI+ePWsvyekL2bPHO7XiIAE
# Xf2FpdQQGqjmKQrTzHeJzQZcR1dm1kghcMHHzEIgZFM7uA0qjefoNSUUlE3bc99+
# ERIc2j8ec6XRfe9/FoAwyIIbGT0RM9oa2bGSMvsK6XiKLNA3Sm2hqhnWyT/21Tbx
# vGOuvbRE5/V1KD+gRZxNpdTcoCFHPadAWNlUp6g6JZLJhXaiRBEEZ4opRHiWDPSl
# Q3suG8OP9f2iyGgC1EjoVPa3Ont7rGVCmN17zwoUDI6PBhOOlVXAlZ+q4NEi3xL8
# 1Bgb9Y4hekP5S03D4fL0ygzJq6w/aCt8IHbF9e9kYxFXx/0HP+xikACAJIK00CYk
# aekZKzCCB54wggWGoAMCAQICEzMAAAAHh6M0o3uljhwAAAAAAAcwDQYJKoZIhvcN
# AQEMBQAwdzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjFIMEYGA1UEAxM/TWljcm9zb2Z0IElkZW50aXR5IFZlcmlmaWNhdGlvbiBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDIwMB4XDTIxMDQwMTIwMDUyMFoX
# DTM2MDQwMTIwMTUyMFowYzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjE0MDIGA1UEAxMrTWljcm9zb2Z0IElEIFZlcmlmaWVkIENv
# ZGUgU2lnbmluZyBQQ0EgMjAyMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBALLwwK8ZiCji3VR6TElsaQhVCbRS/3pK+MHrJSj3Zxd3KU3rlfL3qrZilYKJ
# NqztA9OQacr1AwoNcHbKBLbsQAhBnIB34zxf52bDpIO3NJlfIaTE/xrweLoQ71lz
# CHkD7A4As1Bs076Iu+mA6cQzsYYH/Cbl1icwQ6C65rU4V9NQhNUwgrx9rGQ//h89
# 0Q8JdjLLw0nV+ayQ2Fbkd242o9kH82RZsH3HEyqjAB5a8+Ae2nPIPc8sZU6ZE7iR
# rRZywRmrKDp5+TcmJX9MRff241UaOBs4NmHOyke8oU1TYrkxh+YeHgfWo5tTgkoS
# MoayqoDpHOLJs+qG8Tvh8SnifW2Jj3+ii11TS8/FGngEaNAWrbyfNrC69oKpRQXY
# 9bGH6jn9NEJv9weFxhTwyvx9OJLXmRGbAUXN1U9nf4lXezky6Uh/cgjkVd6CGUAf
# 0K+Jw+GE/5VpIVbcNr9rNE50Sbmy/4RTCEGvOq3GhjITbCa4crCzTTHgYYjHs1Nb
# Oc6brH+eKpWLtr+bGecy9CrwQyx7S/BfYJ+ozst7+yZtG2wR461uckFu0t+gCwLd
# N0A6cFtSRtR8bvxVFyWwTtgMMFRuBa3vmUOTnfKLsLefRaQcVTgRnzeLzdpt32cd
# YKp+dhr2ogc+qM6K4CBI5/j4VFyC4QFeUP2YAidLtvpXRRo3AgMBAAGjggI1MIIC
# MTAOBgNVHQ8BAf8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFNlB
# KbAPD2Ns72nX9c0pnqRIajDmMFQGA1UdIARNMEswSQYEVR0gADBBMD8GCCsGAQUF
# BwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3Np
# dG9yeS5odG0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwDwYDVR0TAQH/BAUw
# AwEB/zAfBgNVHSMEGDAWgBTIftJqhSobyhmYBAcnz1AQT2ioojCBhAYDVR0fBH0w
# ezB5oHegdYZzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWlj
# cm9zb2Z0JTIwSWRlbnRpdHklMjBWZXJpZmljYXRpb24lMjBSb290JTIwQ2VydGlm
# aWNhdGUlMjBBdXRob3JpdHklMjAyMDIwLmNybDCBwwYIKwYBBQUHAQEEgbYwgbMw
# gYEGCCsGAQUFBzAChnVodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2Nl
# cnRzL01pY3Jvc29mdCUyMElkZW50aXR5JTIwVmVyaWZpY2F0aW9uJTIwUm9vdCUy
# MENlcnRpZmljYXRlJTIwQXV0aG9yaXR5JTIwMjAyMC5jcnQwLQYIKwYBBQUHMAGG
# IWh0dHA6Ly9vbmVvY3NwLm1pY3Jvc29mdC5jb20vb2NzcDANBgkqhkiG9w0BAQwF
# AAOCAgEAfyUqnv7Uq+rdZgrbVyNMul5skONbhls5fccPlmIbzi+OwVdPQ4H55v7V
# OInnmezQEeW4LqK0wja+fBznANbXLB0KrdMCbHQpbLvG6UA/Xv2pfpVIE1CRFfNF
# 4XKO8XYEa3oW8oVH+KZHgIQRIwAbyFKQ9iyj4aOWeAzwk+f9E5StNp5T8FG7/VEU
# RIVWArbAzPt9ThVN3w1fAZkF7+YU9kbq1bCR2YD+MtunSQ1Rft6XG7b4e0ejRA7m
# B2IoX5hNh3UEauY0byxNRG+fT2MCEhQl9g2i2fs6VOG19CNep7SquKaBjhWmirYy
# ANb0RJSLWjinMLXNOAga10n8i9jqeprzSMU5ODmrMCJE12xS/NWShg/tuLjAsKP6
# SzYZ+1Ry358ZTFcx0FS/mx2vSoU8s8HRvy+rnXqyUJ9HBqS0DErVLjQwK8VtsBde
# kBmdTbQVoCgPCqr+PDPB3xajYnzevs7eidBsM71PINK2BoE2UfMwxCCX3mccFgx6
# UsQeRSdVVVNSyALQe6PT12418xon2iDGE81OGCreLzDcMAZnrUAx4XQLUz6ZTl65
# yPUiOh3k7Yww94lDf+8oG2oZmDh5O1Qe38E+M3vhKwmzIeoB1dVLlz4i3IpaDcR+
# iuGjH2TdaC1ZOmBXiCRKJLj4DT2uhJ04ji+tHD6n58vhavFIrmcxghcyMIIXLgIB
# ATBxMFoxCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xKzApBgNVBAMTIk1pY3Jvc29mdCBJRCBWZXJpZmllZCBDUyBFT0MgQ0EgMDQC
# EzMABgooJdKxBqM2HEUAAAAGCigwDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3
# AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQx
# IgQg8PX9PEzuPH5wvxhjZd32luMu1atlrGiEyyvFESM6w6swDQYJKoZIhvcNAQEB
# BQAEggGAaWkdluf8LNxYnY2T50A8gBeSLyqwYem/JmNkefkNkzRA1LUZZTJei2JH
# 0/L88bJPdG+VfiWq49WnQ7+0fVz0hV9kb4BUMA3GbgCQ0HBW7MOXCLemOZE8nvL1
# fjzIC3wxQtNLZRroazuzvpncvZMj0pWXAW0ht0JlXLoWIIlym0Le9uCFAf1p691t
# ZCXTmJUelIN7sFRJ/2nc5H5MjzJFolHUn2WQHx+IZ92sxUgO99EfuUYQNKjZ+jyf
# 4qn1fHE3xFlaRnyAQsuMIgWoglo02dUezE2rcEfCovWxWUN1oAkNhVX+aJ0H1ZYa
# VT49WWk3hH3rfnbuztrn02YUbu+SOgxtCozLh85Cu+ja5qWLC0bJQhIlb7GOQpwG
# mQD7AAuHOpVtk7CfsAfBnjHg/mWipNoCPvLRnitBkQwU/LDPfK6O57PNJd7c+TQf
# TZCGXQxEwWTecCGJx7PhQagKqqUN02Gz80MnNWttwSsJJbEkgBHWQbj/QT5eMn26
# 4XNR0jNAoYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIU
# izCCFIcCAQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSC
# AVUwggFRAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIIh//mPV+TKn
# A/sr658eINiEDsYoe8MC+t8rTZdpHaUKAgZqmrFJS4kYEzIwMjYwOTE0MTYwODIw
# LjE1N1owBIACAfSggemkgeYwgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMg
# TGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjdCMUEtMDVFMC1EOTQ3
# MTUwMwYDVQQDEyxNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1
# dGhvcml0eaCCDykwggeCMIIFaqADAgECAhMzAAAABeXPD/9mLsmHAAAAAAAFMA0G
# CSqGSIb3DQEBDAUAMHcxCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xSDBGBgNVBAMTP01pY3Jvc29mdCBJZGVudGl0eSBWZXJpZmlj
# YXRpb24gUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAyMDAeFw0yMDExMTky
# MDMyMzFaFw0zNTExMTkyMDQyMzFaMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMg
# UlNBIFRpbWVzdGFtcGluZyBDQSAyMDIwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAnnznUmP94MWfBX1jtQYioxwe1+eXM9ETBb1lRkd3kcFdcG9/sqtD
# lwxKoVIcaqDb+omFio5DHC4RBcbyQHjXCwMk/l3TOYtgoBjxnG/eViS4sOx8y4gS
# q8Zg49REAf5huXhIkQRKe3Qxs8Sgp02KHAznEa/Ssah8nWo5hJM1xznkRsFPu6rf
# DHeZeG1Wa1wISvlkpOQooTULFm809Z0ZYlQ8Lp7i5F9YciFlyAKwn6yjN/kR4fkq
# uUWfGmMopNq/B8U/pdoZkZZQbxNlqJOiBGgCWpx69uKqKhTPVi3gVErnc/qi+dR8
# A2MiAz0kN0nh7SqINGbmw5OIRC0EsZ31WF3Uxp3GgZwetEKxLms73KG/Z+MkeuaV
# DQQheangOEMGJ4pQZH55ngI0Tdy1bi69INBV5Kn2HVJo9XxRYR/JPGAaM6xGl57E
# i95HUw9NV/uC3yFjrhc087qLJQawSC3xzY/EXzsT4I7sDbxOmM2rl4uKK6eEpurR
# duOQ2hTkmG1hSuWYBunFGNv21Kt4N20AKmbeuSnGnsBCd2cjRKG79+TX+sTehawO
# oxfeOO/jR7wo3liwkGdzPJYHgnJ54UxbckF914AqHOiEV7xTnD1a69w/UTxwjEug
# pIPMIIE67SFZ2PMo27xjlLAHWW3l1CEAFjLNHd3EQ79PUr8FUXetXr0CAwEAAaOC
# AhswggIXMA4GA1UdDwEB/wQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4E
# FgQUa2koOjUvSGNAz3vYr0npPtk92yEwVAYDVR0gBE0wSzBJBgRVHSAAMEEwPwYI
# KwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9S
# ZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIE
# DB4KAFMAdQBiAEMAQTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFMh+0mqF
# KhvKGZgEByfPUBBPaKiiMIGEBgNVHR8EfTB7MHmgd6B1hnNodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBJZGVudGl0eSUyMFZl
# cmlmaWNhdGlvbiUyMFJvb3QlMjBDZXJ0aWZpY2F0ZSUyMEF1dGhvcml0eSUyMDIw
# MjAuY3JsMIGUBggrBgEFBQcBAQSBhzCBhDCBgQYIKwYBBQUHMAKGdWh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwSWRlbnRp
# dHklMjBWZXJpZmljYXRpb24lMjBSb290JTIwQ2VydGlmaWNhdGUlMjBBdXRob3Jp
# dHklMjAyMDIwLmNydDANBgkqhkiG9w0BAQwFAAOCAgEAX4h2x35ttVoVdedMeGj6
# TuHYRJklFaW4sTQ5r+k77iB79cSLNe+GzRjv4pVjJviceW6AF6ycWoEYR0LYhaa0
# ozJLU5Yi+LCmcrdovkl53DNt4EXs87KDogYb9eGEndSpZ5ZM74LNvVzY0/nPISHz
# 0Xva71QjD4h+8z2XMOZzY7YQ0Psw+etyNZ1CesufU211rLslLKsO8F2aBs2cIo1k
# +aHOhrw9xw6JCWONNboZ497mwYW5EfN0W3zL5s3ad4Xtm7yFM7Ujrhc0aqy3xL7D
# 5FR2J7x9cLWMq7eb0oYioXhqV2tgFqbKHeDick+P8tHYIFovIP7YG4ZkJWag1H91
# KlELGWi3SLv10o4KGag42pswjybTi4toQcC/irAodDW8HNtX+cbz0sMptFJK+KOb
# AnDFHEsukxD+7jFfEV9Hh/+CSxKRsmnuiovCWIOb+H7DRon9TlxydiFhvu88o0w3
# 5JkNbJxTk4MhF/KgaXn0GxdH8elEa2Imq45gaa8D+mTm8LWVydt4ytxYP/bqjN49
# D9NZ81coE6aQWm88TwIf4R4YZbOpMKN0CyejaPNN41LGXHeCUMYmBx3PkP8ADHD1
# J2Cr/6tjuOOCztfp+o9Nc+ZoIAkpUcA/X2gSMkgHAPUvIdtoSAHEUKiBhI6JQivR
# epyvWcl+JYbYbBh7pmgAXVswggefMIIFh6ADAgECAhMzAAAAWXzacemNXvXAAAAA
# AABZMA0GCSqGSIb3DQEBDAUAMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMgUlNB
# IFRpbWVzdGFtcGluZyBDQSAyMDIwMB4XDTI2MDEwODE4NTkwMVoXDTI3MDEwNzE4
# NTkwMVowgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTAr
# BgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUG
# A1UECxMeblNoaWVsZCBUU1MgRVNOOjdCMUEtMDVFMC1EOTQ3MTUwMwYDVQQDEyxN
# aWNyb3NvZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1dGhvcml0eTCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKYu5/40eEX+hT+5jFa146bid3dA
# 4LnXYntvkP3CGw4LGARFhnvLMSJ/VtsubzDaeFnm7yb2KSM70WmHQprdCVqpvUH7
# l0uB4jNw7urLoAR9kKHLE0VlMlDStDSxUBI3qwsdrjvdmvV0k+9/njuDEiSlzJTf
# 7Dowd1K3bO4beRyaFhR+Y8tymECOqlOAffYrG2wZdVM51+QSBSe+PEykr8C6Onnq
# SipuF8fZvCb6/huk0Zm6ZwsaixSHIAT2IEGvS7c63Im8jV3a8R0K6i2yiw0NNlnT
# Spwy/Zfv7iwsLBwhfbjBTn+XOl6mPzDXQQ3V+SRP9xXbGKOsBTxzGid7aKAHw3o4
# Ahl9UGWLH9kNP3VUokE6JYkjlfpuUGZ6gQyqDewfxD4VoYIlopt4HZ0xQvqajuJx
# +cr8LR/IZ56gLLmwyMzde5+vtjBoilry/gSZwVGwgkvkIgpKPBQHGsSB0y3szr7Y
# 7wEb6v0yZal1XUvWnnz3inTaSWsCFrLPVwVmXy3ncY5/d25VpOkht+m697GWNbvs
# NOhAOHRaftE9j/hhkoM6RsyJfBLnhqMcA/wcavf5oj5NeyRQdGZeLKcls9csKS3s
# BUzPidxx2iiNH9CPaDq/bLJEOXasYohXMnRinu+fUk81s8VO7DQSF6ffn5oqSHoV
# 8lf1Ax6u+kdShb8BAgMBAAGjggHLMIIBxzAdBgNVHQ4EFgQUj5bnC18D0vlnSRhC
# OiODGGuXNnYwHwYDVR0jBBgwFoAUa2koOjUvSGNAz3vYr0npPtk92yEwbAYDVR0f
# BGUwYzBhoF+gXYZbaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwv
# TWljcm9zb2Z0JTIwUHVibGljJTIwUlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAy
# MDIwLmNybDB5BggrBgEFBQcBAQRtMGswaQYIKwYBBQUHMAKGXWh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwUHVibGljJTIw
# UlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNydDAMBgNVHRMBAf8EAjAA
# MBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDBmBgNVHSAE
# XzBdMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQC
# MA0GCSqGSIb3DQEBDAUAA4ICAQBEMhzC/ZcjpG/zURE7z2Yp5vrUxUjsE5Xa3t/2
# RGvESwvbmsk3bLHhSFAajgo2XQ8xoGDP3sUhKCLPeICSbkVv6V8sSp8fJ8Jos6yr
# awf2YVis8tcV+OO7U9S6JGPQzpmPncfzQc4ne1fqZ4+HiKabIDEoFdddQT2Egkk9
# fzxCY/EZ52avJ27dSfrI/IDmyn9V10O3iQpg2F+C9vNTrk7nVgoDoHa9+Q3pYr0I
# HGnSmt5irgGT436zo5WnXP8FxMhswH1aiyiSZiVzhor10C9C52cP3C8/PEoMKUXs
# tLjoPO0TMkeW/1Fr186KXD45QRgBo0xImgtWTdzWFnlD+p7+iDBIuSrNcRXDRYuq
# /aYZaDhWSI0SYdPIWVh5XvXuWA31a8oQ0SO+oPa3Nk80k0864wiiyJ1KsbSnaaef
# g9vspeghrpY8ljCwxfCUtx5HQRNgAJOI8IKACK4d014Mk0hlRO0lQVRHegqIg29K
# 6Xqkc360W2ZJGUcstlKokkVj6KAHjGyrLRPzepYfiZUJq4gXyxbpvKb1XJ2FN268
# 2aUoNXo9RyRK1ch0f66k6+yj88kzvuC7+vJWtNDs/UpIM6Hhm0kU64JUJ7MMEQcA
# c7kpft7Gm7YeRK+oKgqUgYXCfmzbX8nJXJZnPa8ADWVsIqsuNAxCI0CZXkULofqo
# 5Be6zzGCA9QwggPQAgEBMHgwYTELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFB1YmxpYyBSU0Eg
# VGltZXN0YW1waW5nIENBIDIwMjACEzMAAABZfNpx6Y1e9cAAAAAAAFkwDQYJYIZI
# AWUDBAIBBQCgggEtMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG
# 9w0BCQQxIgQgp8pZqKHcu+EwsGlfwuuYYwVgqD/zb60R2OjPAX6k+cwwgd0GCyqG
# SIb3DQEJEAIvMYHNMIHKMIHHMIGgBCDLRbqx24bpscXEJ+Hjj9xrcUVw7R8OyyMf
# SB2YGK3+vDB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1l
# c3RhbXBpbmcgQ0EgMjAyMAITMwAAAFl82nHpjV71wAAAAAAAWTAiBCBTYfgEGroX
# gyBuem7f4MA7252lQsmYk3pdpb1HxJ7iPjANBgkqhkiG9w0BAQsFAASCAgBVHxj0
# /rM0oDulvppRgPTdO4BjM/AOFYJsiW7/Tygk0CTeFNmHAif+SLvAG755ddNcI0Ah
# nVQclzyPRriyCPAG2VLsPCfZJoS5Wa1DPoDingPxtWC3o9sv+SEC4TEByvOS7Bf5
# 3S2U6j/4ImnMpSZsetCiza6RZbMdILlWMNirGfp7OKhZvd93QEh4Ybo95XwDeEb4
# bDDWtVzSFz/N3E3jhFrDoB7JGubGP8s8Hf3bRhhiD0DPHCc6RfVShQso3aidZ1cq
# 292Es2VElkU/9oRa1fNkPQVEarhHHJujm0gDjgXjINYKxYi/dmkKfjTNhKFoxcC3
# 6AKsFe6cIl6gogDAxwGBiBa/Th0XwmHSrVOk5OUfJ8sMAu/wSFnTY/kXc2yuU4O+
# 0MYArvFxnv7pgKF9DXV+P5N46wBOTxJA0IEFSteCdxputYKv0xQzS8dsSUnpgrz9
# 1/rmrWLbG+6lbBP50B/0LP1Op74AMXj0Qt8Fr9CA9rJzW0GSnV6XEJ/mVt1xIGZA
# Vgk0gN2uY/OIoTsNOkYx4GafQwccVq/UCHrmMHg+3TGJYKp1CkHpAdteS5U/llbW
# YKe84HrvO46WNh8GK8KIH9ZdHai90Ziz0oovysl4LWzZZB3f5EFrDdpPjg1dnsOr
# LxnrxIKdRRnX/0iKGICewQtovCQVlpYiilqmYw==
# SIG # End signature block
