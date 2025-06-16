# Version: v1.0.33

# TODO: 
# * DONE: Add RC to GRIPSDirectPrint
#       - Pass to WS when adding updating printers
#       - filter on RC when displaying printers page    
# * DONE: Pick up default printer and use when no printer is specified
# * DONE: Adapt all places where printers are displayed/selected and show normal printers page instead of Web Client Printers when Printing App is disabled
# * DONE: Renumber objects according to GRIPS rules - Viktoras says its OK to use 90000+
# * DONE: Move Settings config.json
# * DONE: Encrypt password using secret protected key read from the registry
# * DONE: Create installation script to install processor as service using nssm
#       - DONE: Should prompt for parameters during installation
#       - DONE: List of URLs with by country - separate file to settings so can be modified centrally
#       - DONE: Register file type for PDF signing
#               cmd.exe /c assoc .signpdf=SignedPDFFile               
#               cmd.exe /c ftype SignedPDFFile="""C:\Program Files\signotec\signoSign2\SignoSign2.exe""" """%1"""
#       - DONE: Create an self-extracting archive with the script and the settings file:
# * DONE: Rotate the transcript file every so often and keep only the last n days
# * DONE: Make overrides for settings in config.json - load userconfig.json if exists
# * DONE: Make self-updating - see info. from chatGPT saved in BC DirectPrinting folder
# * DONE: Handle additional arguments e.g. "-sign" for Signosign (create field on GRIPSDirectPrintQueue table and fill from printer selection using events)
# *     - Handled as a download so that software can option in user's session
# * TODO: Localize the strings - see grips.net\WebServerSidePrinter\NAVPrintingApplication\GRIPSWebPrintingApplication-Install.ps1

param (
    [string]$configFile = "$PSScriptRoot\config.json",
    [string]$userConfigFile = "$PSScriptRoot\userconfig.json"
)

### POWERSHELL ON WINDOWS ###

# Function to get the decrypted credentials from the encrypted file
function Get-StoredCredential {
    param([string]$credFile,
        $key)

    if (Test-Path -Path $credFile -PathType Leaf) {
        $credArray = Get-Content $credFile
        $credential = New-Object -TypeName System.Management.Automation.PSCredential `
            -ArgumentList $credArray[0], ($credArray[1] | ConvertTo-SecureString -Key $key)
        return $credential
    }
}

# Load configuration from JSON file
$config = Get-Content $configFile | ConvertFrom-Json

# Check if userconfig.json exists
if (Test-Path -Path $userConfigFile -PathType Leaf) {
    # Load user configuration from userconfig.json
    $userConfig = Get-Content $userConfigFile | ConvertFrom-Json

    # Update or add keys from user configuration
    $userConfig.PSObject.Properties | ForEach-Object {
        $config.$($_.Name) = $_.Value
    }
}

$releaseApiUrl = $config.ReleaseApiUrl;

$keyPath = "$PSScriptRoot\Installer\l02fKiUY\l02fKiUY.txt"
$key = @(((Get-Content $keyPath) -split ","))

$credFile = "$PSScriptRoot\$($config.BasicAuthLogin).TXT"

$credential = Get-StoredCredential -credFile $credFile -key $key

# Authentication:
$Authentication = @{
    #"Company"                     = 'NAS Company' # Note: Must exist or be left empty if a Default Company is setup in the Service Tier. Only used for authentication as printers and jobs are PerCompany=false
    "Company"                     = $config.Company

    "BasicAuthLogin"              = $config.BasicAuthLogin;
    "BasicAuthPassword"           = $(([Net.NetworkCredential]::new('', $credential.Password).Password))

    "OAuth2CustomerAADIDOrDomain" = $config.OAuth2CustomerAADIDOrDomain
    "OAuth2ClientID"              = $config.OAuth2ClientID
    "OAuth2ClientSecret"          = $config.OAuth2ClientSecret
}
#

### Configuration ###

# URLs for webservices:
#$BaseURL    = "https://<hostname>/<instance>/ODataV4/"
$BaseURL = $config.BaseURL
$RespCtr = $config.RespCtr

$PrintersWS = "GRIPSDirectPrintPrinterWS"
$QueuesWS = "GRIPSDirectPrintQueueWS"

# Misc.:
#$IgnorePrinters = @("OneNote for Windows 10","Microsoft XPS Document Writer","Microsoft Print to PDF","Fax") # Don't offer these printers to Business Central
$IgnorePrinters = $config.IgnorePrinters

#$PDFPrinter_exe  = "$PSScriptRoot\PDFXCview\PDFXCview.exe"
if (-not [System.IO.Path]::IsPathRooted($config.PDFPrinter_exe)) {
    $PDFPrinter_exe = "$PSScriptRoot\$($config.PDFPrinter_exe)"
}
else {
    $PDFPrinter_exe = $config.PDFPrinter_exe
}

$Sign_exe = $config.Sign_exe
$Sign_params = $config.Sign_params

# {0} = PrinterName
# {1} = FileName
# {2} = Papersource Argument e.g. bin=257,
# {3} = Additional Arguments
#$PDFPrinter_params = "/printto ""{0}"" ""{1}""" # PDFXCview 
#$PaperSourceArgument = "" #PDFXCview

#$PDFPrinter_params = "-print-to ""{0}"" -print-settings ""{2}{3}"" ""{1}""" # SumatraPDF
$PDFPrinter_params = $config.PDFPrinter_params
$PaperSourceArgument = "bin={0}," # SumatraPDF

#$Delay = 2 # Delay between checking for print jobs in seconds
$Delay = $config.Delay

#$UpdateDelay = 300 # Delay between updating printers in seconds
$UpdateDelay = $config.UpdateDelay

#$ReleaseCheckDelay = 600 # Delay between checking for new releases in seconds
$ReleaseCheckDelay = $config.ReleaseCheckDelay

### End of Configuration ###

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
    $transcripts = Get-ChildItem -Path $transcriptPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$config.TranscriptMaxAgeDays) }
    $transcripts | Remove-Item -Force

    return [datetime]::Now
}


function Get-BasicAuthentication {
    param (
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Login,
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Password
    )
    PROCESS { 
        return [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$Login`:$Password"))
    }
}

function Get-OAuth2AccessToken {
    param (
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$ClientID,
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$ClientSecret,
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$CustomerAAD_ID_Or_Domain
    )
    PROCESS { 
        Add-Type -AssemblyName System.Web
        $Body = "client_id=" + [System.Web.HttpUtility]::UrlEncode($ClientID) + "&client_secret=" + [System.Web.HttpUtility]::UrlEncode($ClientSecret) +
        "&scope=https://api.businesscentral.dynamics.com/.default&grant_type=client_credentials"
        Try {
            $Json = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$CustomerAAD_ID_Or_Domain/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $Body 
        }
        Catch {
            $Reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $Reader.BaseStream.Position = 0
            $Reader.DiscardBufferedData()
            Write-Host ($Reader.ReadToEnd() | ConvertFrom-Json).error.message -ForegroundColor Red
        }

        return $Json.access_token
    }
}

function Invoke-BCWebService {
    param (
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Method,
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$BaseURL,
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$WebServiceName,
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$DirectLookup,
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$Filter,
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$ETag,
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Object]$Authentication,
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$Body,
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch]$GetParametersOnly
    )
    PROCESS { 
        $URL = $BaseURL.trimend("/")

        $Headers = @{"Accept" = "application/json" }
        if (($Authentication.BasicAuthLogin -ne "") -and ($Authentication.BasicAuthPassword -ne "")) {
            $Headers.Add("Authorization", "Basic $(Get-BasicAuthentication -Login $Authentication.BasicAuthLogin -Password $Authentication.BasicAuthPassword)")
        }
        else {
            $Headers.Add("Authorization", "Bearer $(Get-OAuth2AccessToken -ClientID $Authentication.OAuth2ClientID -ClientSecret $Authentication.OAuth2ClientSecret `
                                                                         -CustomerAAD_ID_Or_Domain $Authentication.OAuth2CustomerAADIDOrDomain)")
        }

        if ($Method -eq "Get") {
            $Headers.Add("Data-Access-Intent", "ReadOnly")
        }

        if (-not [string]::IsNullOrEmpty($Body)) {
            $Headers.Add("Content-Type", "application/json")
        }
        
        if (-not [string]::IsNullOrEmpty($ETag)) {
            $Headers.Add("If-Match", $ETag)
        }
        
        if (-not ([string]::IsNullOrEmpty($Authentication.Company))) {
            $URL = "$URL/Company('$($Authentication.Company)')"
        }

        $URL = "$URL/$WebServiceName"

        if (-not ([string]::IsNullOrEmpty($DirectLookup))) {
            $URL = "$URL($DirectLookup)"
        }

        if (-not ([string]::IsNullOrEmpty($Filter))) {
            $URL = "$URL`?`$filter=$Filter"
        }

        $Parameters = @{
            Method  = $Method
            Uri     = $URL
            Headers = $Headers
        }

        if (-not [string]::IsNullOrEmpty($Body)) {
            $Parameters.Add("Body", $Body)
        }

        if ($GetParametersOnly) {
            return $Parameters
        }
        else {
            Try {
                $Response = Invoke-RestMethod @Parameters
            }
            Catch { 
                $Reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $Reader.BaseStream.Position = 0
                $Reader.DiscardBufferedData()
                $Response = $Reader.ReadToEnd()
                Write-Host "Error calling $($Parameters.Values): $Response" -ForegroundColor Red
                Write-Host ($Response | ConvertFrom-Json).error.message -ForegroundColor Red
            }

            return $Response
        }
    }
}

function SafePrinterName {
    param (
        [String]$PrinterName
    )

    return $PrinterName -replace "\\", "``"
}

function RealPrinterName {
    param (
        [String]$PrinterName
    )

    return $PrinterName -replace "``", "`\"
}

function Update-Check {
    Write-Host "Checking for updates..." -ForegroundColor White
    
    # Start background job to check for updates
    Start-Job -Arg $releaseApiUrl, $currentVersion, $updateSignalFile -ScriptBlock {
        param (
            [string]$releaseApiUrl,
            [string]$currentVersion,
            [string]$updateSignalFile
        )

        $LatestRelease = Invoke-RestMethod -Uri $releaseApiUrl -Method Get
        $releaseVersion = $LatestRelease.tag_name.TrimStart('v')
        # Compare versions
        if ([version]$releaseVersion -gt [version]$currentVersion) {
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
        
            # Signal the main script that the update is ready
            Set-Content -Path $updateSignalFile -Value "$($extractedSubFolder.FullName)"

            # Clean up temporary files
            Remove-Item -Path $TempZipFile -Force
        }
        else {
            Write-Output "No update required. Current version ($global:currentVersion) is up to date."
        }
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

    Get-ScriptVersion -ScriptPath $FullScriptPath
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Script updated to version $global:currentVersion."

    # Exit the script with non-zero exit code to force the service to restart
    Exit 1
}

function Get-ScriptVersion {
    param (
        [string]$ScriptPath
    )

    $scriptContent = Get-Content -Path $ScriptPath
    # Extract the current version from the script
    $global:currentVersion = $null
    foreach ($line in $scriptContent) {
        if ($line -match "#\s*Version:\s*v?(\d+\.\d+\.\d+)") {
            $global:currentVersion = $Matches[1]
            break
        }
    }        
    Write-Output "Script version: $global:currentVersion"
}

#House keeping
# Get the full path of the directory containing the script
$ScriptPath = $PSScriptRoot

# Get the filename of the script
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)

$LastTranscriptRotation = Start-MyTranscript 

$ErrorActionPreference = 'Continue'
$LastPrinterUpdate = (Get-Date).AddSeconds(-$UpdateDelay) # Make sure the update is run immediately on startup of the script
$LastReleaseCheck = (Get-Date).AddSeconds(-$ReleaseCheckDelay) # Make sure the release check is run immediately on startup of the script

# To combine them into the full path to the script file
$FullScriptPath = Join-Path -Path $ScriptPath -ChildPath $ScriptName

Write-Host "Starting up - current script version: $(Get-ScriptVersion -ScriptPath $FullScriptPath)"

# Define update signal file
$updateSignalFile = "$ScriptPath\update_ready.txt"

while ($true) {
    # Check for new releases
    if (($(Get-Date) - $LastReleaseCheck).TotalSeconds -gt $ReleaseCheckDelay) {
        Update-Check
        $LastReleaseCheck = Get-Date
    }

    if (Test-Path -Path $updateSignalFile) {
        Update-Release
    }

    # Rotate the transcript file when needed
    if ((Get-Date) -gt $LastTranscriptRotation.AddMinutes($config.TranscriptRotationTimeMins)) {
        Stop-Transcript
        $LastTranscriptRotation = Start-MyTranscript
    }

    #Fetch printers on this host from BC    
    Clear-Variable -Name "BCPrinters" -ErrorAction SilentlyContinue
    $BCPrinters = (Invoke-BCWebService -Method Get -BaseURL $BaseURL -WebServiceName $PrintersWS -Filter "HostID eq '$env:COMPUTERNAME'" -Authentication $Authentication).value

    #Register new printers in BC
    foreach ($Printer in (Get-Printer | Where-Object { $IgnorePrinters -notcontains $_.Name } | Where-Object { $BCPrinters.PrinterID -notcontains $(SafePrinterName($_.Name)) })) {
        $defaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
        if ($Printer.Name -eq $defaultPrinter.Name) {
            $isDefault = 'true'
        }
        else {
            $isDefault = 'false'
        }
        Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Adding new printer in BC: $($Printer.Name)" -ForegroundColor Yellow
        Invoke-BCWebService -Method Post -BaseURL $BaseURL -WebServiceName $PrintersWS -Authentication $Authentication `
            -Body "{""HostID"":""$($env:COMPUTERNAME)"",""PrinterID"":""$(SafePrinterName($Printer.Name))"",""ResponsibilityCenter"":""$($RespCtr)"",""DefaultPrinter"":""$isDefault""}" | Out-Null
    }
          
    #Update existing printers in BC
    if (($(Get-Date) - $LastPrinterUpdate).TotalSeconds -gt $UpdateDelay) {
        $defaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }

        foreach ($Printer in (Get-Printer | Where-Object { $IgnorePrinters -notcontains $_.Name } | Where-Object { $BCPrinters.PrinterID -contains $(SafePrinterName($_.Name)) })) {
            $BCPrinter = $BCPrinters | Where-Object { $(SafePrinterName($Printer.Name)) -eq $_.PrinterID }
            if ($Printer.Name -eq $defaultPrinter.Name) {
                $isDefault = 'true'
            }
            else {
                $isDefault = 'false'
            }
            Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Updating printer in BC (RowNo: $($BCPrinter.RowNo)): $(SafePrinterName($Printer.Name))" -ForegroundColor Yellow   
            Invoke-BCWebService -Method Patch -BaseURL $BaseURL -WebServiceName $PrintersWS -DirectLookup $BCPrinter.RowNo -ETag $BCPrinter."@odata.etag" -Authentication $Authentication `
                -Body "{""HostID"":""$env:COMPUTERNAME"",""PrinterID"":""$(SafePrinterName($Printer.Name))"",""ResponsibilityCenter"":""$RespCtr"",""DefaultPrinter"":""$isDefault""}" | Out-Null
        }
        $LastPrinterUpdate = Get-Date
        Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Looking for print jobs every $Delay seconds, updating printers every $UpdateDelay seconds..." -ForegroundColor White   
    }

    #Print the queued jobs for the printers on this host
    if (($BCPrinters.NoQueued | Measure-Object -Sum).Sum -gt 0 ) {
        foreach ($Job in (Invoke-BCWebService -Method Get -BaseURL $BaseURL -WebServiceName $QueuesWS -Filter "HostID eq '$env:COMPUTERNAME' and Status eq 'Queued'" -Authentication $Authentication).value) {
            $Job = Invoke-BCWebService -Method Patch -BaseURL $BaseURL -WebServiceName $QueuesWS -DirectLookup ($Job.RowNo) -ETag ($Job."@odata.etag") `
                -Body "{""Status"":""Printing""}" -Authentication $Authentication
            Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Printing job (RowNo: $($Job.RowNo)) on printer $(RealPrinterName($Job.PrinterID))..." -ForegroundColor Yellow

            $TempFile = New-TemporaryFile
            $PDFFileName = $TempFile.FullName + ".grdp.pdf"
            Remove-Item -Path $TempFile.FullName

            if ($Job.AddArgs -ieq "-sign") {
                $Action = "signing"
                $Executable = $Sign_exe
                $Params = $Sign_params -f $PDFFileName
            }
            else {
                $Action = "printing"
                $Executable = $PDFPrinter_exe

                $Papersource = $Job.RawKind
                $AddArgs = $Job.AddArgs

                if ([string]::IsNullOrEmpty($Papersource)) {
                    $PaperSourceArgument = ""
                }
                else {
                    $PaperSourceArgument = $PaperSourceArgument -f $Papersource
                }

                if ($AddArgs.Contains("-print-settings") -and $Params.Contains("-print-settings")) {
                    $addPrintArgs = $AddArgs -split '\s+'

                    if ($addPrintArgs.Length -gt 1) {
                        $AddArgs = $addPrintArgs[1].Trim('"')
                    }
                }
                
                $Params = $PDFPrinter_params -f $($Job.PrinterID -replace "``", "`\"), $PDFFileName, $PaperSourceArgument, $AddArgs
            }

            $InvokeRestMethodParameters = (Invoke-BCWebService -Method Patch -BaseURL $BaseURL -WebServiceName $QueuesWS -DirectLookup ($Job.RowNo) -ETag ($Job."@odata.etag") `
                    -Body "{""Status"":""Printed"",""PrinterMessage"":""Passed to $(Split-Path -Path $Executable -Leaf) for $Action""}" `
                    -Authentication $Authentication -GetParametersOnly)

            Start-Job -Arg $Job, $Executable, $Params, $PaperSourceArgument, $InvokeRestMethodParameters, $PDFFileName -ScriptBlock {
                Param($Job, $Executable, $Params, $PaperSourceArgument, $InvokeRestMethodParameters, $PDFFileName)
                #Start-Transcript -Path "$env:TEMP\GRIPSDirectPrintProcessor_$($Job.RowNo).log" -Append
                Invoke-RestMethod @InvokeRestMethodParameters
                    
                [IO.File]::WriteAllBytes($PDFFileName, [System.Convert]::FromBase64String($Job.PDFPrintJobBASE64))

                Start-Process -FilePath $Executable -ArgumentList $Params -Wait -PassThru
                Remove-Item -Force -Path $PDFFileName
                #Stop-Transcript
            } #| Out-Null

        }
        Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Looking for print jobs every $Delay seconds, updating printers every $UpdateDelay seconds..." -ForegroundColor White   
    }

    Start-Sleep -Seconds $Delay
}

Stop-Transcript

Exit 1 # Exit with non-zero exit code to force the service to restart

<# IF THE SCRIPT HAS BEEN CHANGED THEN IT WILL NEED RESIGNING:
.\CreateSignedScript.ps1 -Path .\Run-GRIPSDirectPrintProcessor.ps1
#>

# SIG # Begin signature block
# MIIP2AYJKoZIhvcNAQcCoIIPyTCCD8UCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7S76PflNqvtZv2MBCrcHGmDX
# rVSggg0oMIIFUzCCBDugAwIBAgITGAAAFn2/inOgnDqjKgAAAAAWfTANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUQFItdBs2O6vgi6iJzZnVz0+zbcYwDQYJKoZI
# hvcNAQEBBQAEggEAQjVUzV9j5DgkPIVfs+YQ7GPcCWI+P6ekLGsO0FZ5bczq+Vjd
# ryUjC3Ue3jUOJt/J6KjdPNqoK9njNL6ZQDZ5Z3uHPaD9/BbY1kVVNntlgSbyRIBO
# cItR2d/mgUQDS84E0zOevQDwLoGIq7rl+vs/PffrgDdWClfb9CWC1mrPzLBG3nLq
# Cxeuz8682/WanKUpi5PVC9eLIO2BeZ6Xup/N13ESw0cArVXpazBde/JtQZGNSjVO
# /z6C2CfuZ2GINXdn8aeBg2jsFXQNLfFSViGRaTcjZLUiioyF9BdykkGFQFASmffY
# anVKpyqMq+WQsb4ihJ/Bh+XJ8W2wy6QjTPnKcA==
# SIG # End signature block
