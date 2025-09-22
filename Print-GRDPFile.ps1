param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    [string]$configFile = ""
)

# Get the full path of the directory containing the script
$ScriptPath = $PSScriptRoot

if ($configFile -eq "") { $configFile = "$ScriptPath\config.json" }
$global:configFile = $configFile

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
    Write-Host "Checking for updates..." -ForegroundColor White
    
    if ($global:config.UsePreleaseVersion) {
        $LatestRelease = (Invoke-RestMethod -Uri ($releaseApiUrl -replace '/latest$', '') -Method Get) | Select-Object -First 1
    }
    else {
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

function Get-ScriptVersion {
    # Load configuration from JSON file
    $global:config = Get-Content $global:configFile | ConvertFrom-Json

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

# Load configuration from JSON file
$config = Get-Content $global:configFile | ConvertFrom-Json

if (-not [System.IO.Path]::IsPathRooted($config.PDFPrinter_exe)) {
    $PDFPrinter_exe = "$ScriptPath\$($config.PDFPrinter_exe)"
}
else {
    $PDFPrinter_exe = $config.PDFPrinter_exe
}

$PDFPrinter_params = $config.PDFPrinter_params
$PaperSourceArgument = "bin={0}," # SumatraPDF
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

            # Find the options text file (assuming only one .txt file)
            $optionsFile = Get-ChildItem -Path $tempFolder -Filter *.txt | Select-Object -First 1

            if (-not $optionsFile) {
                Write-Error "Options text file not found inside the .grdp archive."
                exit 1
            }

            # Parse options
            $options = Get-Options -FilePath $optionsFile.FullName

            # Validate required options
            if (-not $options.ContainsKey("PDFFile")) {
                Write-Error "PDFFile option missing in options file."
                exit 1
            }

            $pdfFilePath = Join-Path -Path $tempFolder -ChildPath $options["PDFFile"]
            if (-not (Test-Path $pdfFilePath)) {
                Write-Error "PDF file specified in options not found: $pdfFilePath"
                exit 1
            }

            # Build SumatraPDF print arguments
            $printerName = $options["Printer"]
            $outputBin = $options["OutputBin"]
            $AddArgs = $options["AdditionalArgs"]

        
            # Construct paper source argument if OutputBin is specified
            if ([string]::IsNullOrEmpty($outputBin)) {
                $PaperSourceArgument = ""
            }
            else {
                $PaperSourceArgument = $PaperSourceArgument -f $outputBin
            }

            if (-not [string]::IsNullOrEmpty($AddArgs)) {
                if ($AddArgs.Contains("-print-settings") -and $Params.Contains("-print-settings")) {
                    $addPrintArgs = $AddArgs -split '\s+'

                    if ($addPrintArgs.Length -gt 1) {
                        $AddArgs = $addPrintArgs[1].Trim('"')
                    }
                } 
            }

            $Params = $PDFPrinter_params -f $printerName, $pdfFilePath, $PaperSourceArgument, $AddArgs

            # Start printing
            Write-Host "Printing $pdfFilePath to printer '$printerName' with settings '$Params'"
            $proc = Start-Process -FilePath $PDFPrinter_exe -ArgumentList $Params -PassThru

            # Wait for process exit with timeout (e.g., 30 seconds)
            if (-not $proc.WaitForExit(30000)) {
                Write-Warning "Print process did not exit within 30 seconds, killing process."
                try {
                    $proc.Kill()
                }
                catch {
                    Write-Warning "Failed to kill print process: $_"
                }
            }
            else {
                Write-Host "Print job completed"
            }        

        }
        finally {
            # Clean up temp folder
            Remove-Item -Path $tempFolder -Recurse -Force
            #Remove-Item -Path $InputFile # tidy up downloaded file
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
    Stop-Transcript
}

<# IF THE SCRIPT HAS BEEN CHANGED THEN IT WILL NEED RESIGNING:
.\CreateSignedScript.ps1 -Path .\Print-GRDPFile.ps1
#>

# SIG # Begin signature block
# MIIP2AYJKoZIhvcNAQcCoIIPyTCCD8UCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfAqQ/r8YqFgPZeuTBNGhKOZf
# lPWggg0oMIIFUzCCBDugAwIBAgITGAAAFn2/inOgnDqjKgAAAAAWfTANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUni1Q4317zvTdP7UQkeN9V+QugEMwDQYJKoZI
# hvcNAQEBBQAEggEAgCnUfSjB70Mz4Jp4G+Bz7GrMtTAfQMe5YcYqEsya1wH2fNXD
# 7DaiiISVFoaVESDfLBQNkVX3rqYkZN4nwO5cojAQN5mYX7y/gskVf9Lw874BgpwG
# tCKQSwvarAo4wlBQOg/U+gfs/T18YBa8ABW1hLDTfsRrkDEvp69Ssa8oEfBNHGuz
# pJNCZtmN2rY6Ayd1r/s+ewxOGndYnj1EGmFoZ8rzq+7lO5TbeDZNvSUMjpm3t9iq
# 7JUWhotZXOQKQVlFvaZqdf4S3jbMY0YoyPqkttWzFilX7z94QtuRVcmLW2KrDlqn
# NO9QTWCeFvUxv7I8yTEShAT7UkdXaxh8qsHf8g==
# SIG # End signature block
