# Version: v1.0.34 #Moved to config.json - can be removed with next update

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
    [string]$configFile = "",
    [string]$userConfigFile = ""
)

### POWERSHELL ON WINDOWS ###

if ($configFile -eq "") { $configFile = "$PSScriptRoot\config.json" }
$global:configFile = $configFile

if ($userConfigFile -eq "") { $userConfigFile = "$PSScriptRoot\userconfig.json" }
$global:userConfigFile = $userConfigFile

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
}

Get-Config -configFile $global:configFile -userConfigFile $global:userConfigFile

$releaseApiUrl = $global:config.ReleaseApiUrl;

$keyPath = "$PSScriptRoot\Installer\l02fKiUY\l02fKiUY.txt"
$key = @(((Get-Content $keyPath) -split ","))

$credFile = "$PSScriptRoot\$($global:config.BasicAuthLogin).TXT"

$credential = Get-StoredCredential -credFile $credFile -key $key

# Authentication:
$Authentication = @{
    #"Company"                     = 'NAS Company' # Note: Must exist or be left empty if a Default Company is setup in the Service Tier. Only used for authentication as printers and jobs are PerCompany=false
    "Company"                     = $global:config.Company

    "BasicAuthLogin"              = $global:config.BasicAuthLogin;
    "BasicAuthPassword"           = $(([Net.NetworkCredential]::new('', $credential.Password).Password))

    "OAuth2CustomerAADIDOrDomain" = $global:config.OAuth2CustomerAADIDOrDomain
    "OAuth2ClientID"              = $global:config.OAuth2ClientID
    "OAuth2ClientSecret"          = $global:config.OAuth2ClientSecret
}
#

### Configuration ###

# URLs for webservices:
#$BaseURL    = "https://<hostname>/<instance>/ODataV4/"
$BaseURL = $global:config.BaseURL
$RespCtr = $global:config.RespCtr

$PrintersWS = "GRIPSDirectPrintPrinterWS"
$QueuesWS = "GRIPSDirectPrintQueueWS"

# Misc.:
#$IgnorePrinters = @("OneNote for Windows 10","Microsoft XPS Document Writer","Microsoft Print to PDF","Fax") # Don't offer these printers to Business Central
$IgnorePrinters = $global:config.IgnorePrinters

#$PDFPrinter_exe  = "$PSScriptRoot\PDFXCview\PDFXCview.exe"
if (-not [System.IO.Path]::IsPathRooted($global:config.PDFPrinter_exe)) {
    $PDFPrinter_exe = "$PSScriptRoot\$($global:config.PDFPrinter_exe)"
}
else {
    $PDFPrinter_exe = $global:config.PDFPrinter_exe
}

$Sign_exe = $global:config.Sign_exe
$Sign_params = $global:config.Sign_params

# {0} = PrinterName
# {1} = FileName
# {2} = Papersource Argument e.g. bin=257,
# {3} = Additional Arguments
#$PDFPrinter_params = "/printto ""{0}"" ""{1}""" # PDFXCview 
#$PaperSourceArgument = "" #PDFXCview

#$PDFPrinter_params = "-print-to ""{0}"" -print-settings ""{2}{3}"" ""{1}""" # SumatraPDF
$PDFPrinter_params = $global:config.PDFPrinter_params
$PaperSourceArgument = "bin={0}," # SumatraPDF

#$Delay = 2 # Delay between checking for print jobs in seconds
$Delay = $global:config.Delay

#$UpdateDelay = 300 # Delay between updating printers in seconds
$UpdateDelay = $global:config.UpdateDelay

#$ReleaseCheckDelay = 600 # Delay between checking for new releases in seconds
$ReleaseCheckDelay = $global:config.ReleaseCheckDelay

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
    $transcripts = Get-ChildItem -Path $transcriptPath | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$global:config.TranscriptMaxAgeDays) }
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
    Start-Job -Arg $releaseApiUrl, $currentVersion, $updateSignalFile, $global:config.UsePrereleaseVersion -ScriptBlock {
        param (
            [string]$releaseApiUrl,
            [string]$currentVersion,
            [string]$updateSignalFile,
            [bool]$usePrereleaseVersion
        )

        if ($usePrereleaseVersion) {
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

    # Exit the script with non-zero exit code to force the service to restart
    Exit 1
}

function Get-ScriptVersion {
    Get-Config -configFile $global:configFile -userConfigFile $global:userConfigFile

    $global:currentVersion = $global:config.Version.TrimStart('v')
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

Write-Host "Starting up - current script version: $(Get-ScriptVersion)"

# Define update signal file
$updateSignalFile = "$ScriptPath\update_ready.txt"

while ($true) {
    # Check if an update is ready to be applied
    if (Test-Path -Path $updateSignalFile) {
        Update-Release
    }
    else {
        # Only check for new releases if no update is pending
        if (($(Get-Date) - $LastReleaseCheck).TotalSeconds -gt $ReleaseCheckDelay) {
            Update-Check
            $LastReleaseCheck = Get-Date
        }
    }

    # Rotate the transcript file when needed
    if ((Get-Date) -gt $LastTranscriptRotation.AddMinutes($global:config.TranscriptRotationTimeMins)) {
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

            Write-Host "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Command: `"$Executable $Params`"" -ForegroundColor Gray

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
# MII6+AYJKoZIhvcNAQcCoII66TCCOuUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDYMGRcSddICfd
# frnidU1Puuw0b2zFMwITGWNYcSo2i6CCIxwwggXMMIIDtKADAgECAhBUmNLR1FsZ
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
# IgQgL8+EvEa3Xe73uGOkZALeDXDFhL74pZxlXJipT6UH5oowDQYJKoZIhvcNAQEB
# BQAEggGAW08HbHw/BTUHQonB0vEhXidl/LVdZW90heVmBR4eQtxsI2yXiREKDNRW
# TPtq7Gng60xsnUXzHQrHwn4dYlIu8U+hZjt0xp06L9Dqh+qWd18/egqrMsGNtlnX
# a3bsDVEj3ZELmwIWHZZVmuvLTluNVtz5djFHiq10dt5aHyNiOtcwULWhbk97tq3P
# 4cdWC+VR0APz0WL2oc3c4cTl4id8IpMrrU5ITUIdOSEXgA+04GQi7y4wje2nQQNd
# +tDraz/muyaam4wxXbG4H5GMZBvh5vZ3aH4W4xtcyah/xOVqraCIi8kW10LFWpQZ
# d4b4Np2yWItYkW9A+BehTNN5NuZAKHvDhyryc8+s5JIcYHmFX0QyWtl1st1XyfzQ
# jH2GMmx3SO0bTEO3we1B6HPsOZI4QJdFs7ereIs/jxY3C+f9NDJkK4SlhLLl27c4
# V7eTG+sFRYqta6iUKUm8kndgkK1RP0HJxytfhCMIaCuAWGcSr7mi2yF02+KXaF2G
# 87BCjML6oYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIU
# izCCFIcCAQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSC
# AVUwggFRAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIGcr1GMYGP5P
# BZH1Qoimu4hC7fwHpd59LMO0o3OshQtOAgZqmsvJl5YYEzIwMjYwOTE0MTYwNzMz
# LjE1N1owBIACAfSggemkgeYwgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMg
# TGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjdBMUEtMDVFMC1EOTQ3
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
# epyvWcl+JYbYbBh7pmgAXVswggefMIIFh6ADAgECAhMzAAAAW0q1jUEybdx0AAAA
# AABbMA0GCSqGSIb3DQEBDAUAMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMgUlNB
# IFRpbWVzdGFtcGluZyBDQSAyMDIwMB4XDTI2MDEwODE4NTkwNVoXDTI3MDEwNzE4
# NTkwNVowgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTAr
# BgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUG
# A1UECxMeblNoaWVsZCBUU1MgRVNOOjdBMUEtMDVFMC1EOTQ3MTUwMwYDVQQDEyxN
# aWNyb3NvZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1dGhvcml0eTCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJBUzBbbnlDXee0B0KD5G4/475th
# FyfctCyuESTWQXvlLi4Wx/td2qUdeq4ideeg6VWhiOHfu3wJV4TUGSRtqh9Ccr1B
# miBKv9iuFpgHyIBu5Qx38ZsxwlFeXVS+ZqJJKnXRbDNQdcYSoC/6c0hQJ/PH50DB
# RDQkPXVwyFizLrRH9AlrJeUg7BKeT23zftS8/KOJLvEEbHOF6pSOY3ZVprZUWbWj
# WwRTmoHaQ/E8vrWtLNyEJ+b089VW1Ikra3t4GTB5Wby3CL1K2zYnAxBIvafsKMFy
# j9OuXHcTPKMDoFSMeamG9MKOMb6uoG1PjdnDgsLP6EOMRSzrLL7jED1mbB9RSd9f
# hty+HQr6vZgsBn6oUy+YTpNVLskwdtUM82WYAkPztlOt3AiL0qyV7/U3j/uq3vHM
# jPM0w0340M57Nei0g4BCcMt0dbqoc91VgCb3/36sHQANontn1HOF2oLk8190QRS4
# 3isHVra8H8sf5+GlqIYsYiCKX04HZiOzZW826nVI6d++8lyTeWmpj90Ua9uPbJhV
# jwE3oh6tO510ySqmSMSLEN07p3Ibe3E6BAb2w93rWzb26+dpSthbKF4kApofqBsW
# PX4MEtHKSOftPmVTCQ47tghrVuHia9jY+Hsj01m4KW4WtkmVm3L6hMZECMa4sjMx
# AXz+bX/AJhWTe6TZAgMBAAGjggHLMIIBxzAdBgNVHQ4EFgQU7/LqUlWWYhXJdXwg
# YKx4b8Gv0rYwHwYDVR0jBBgwFoAUa2koOjUvSGNAz3vYr0npPtk92yEwbAYDVR0f
# BGUwYzBhoF+gXYZbaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwv
# TWljcm9zb2Z0JTIwUHVibGljJTIwUlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAy
# MDIwLmNybDB5BggrBgEFBQcBAQRtMGswaQYIKwYBBQUHMAKGXWh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwUHVibGljJTIw
# UlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNydDAMBgNVHRMBAf8EAjAA
# MBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDBmBgNVHSAE
# XzBdMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQC
# MA0GCSqGSIb3DQEBDAUAA4ICAQAAH+zd+XKh4OxXYMWFmtgilXAQGctOjCUB1w/u
# BiC/OXcH3Ia4/XbdUhKzFbaiTbIE6vYZKd1p4u7nKOLkawymAMVyuO7LSl6rLKtt
# ZIyLhWjTK0zXOz0u4xLq9+bRtBEKJvA6sD5nJwH1IO6z1YizyuIRoalMCnbrUixf
# WxQn4TAmN7t9uk+X2FUThEa3ewzRwhtG+xwaAbLMkxRmR24JnfXd1VxKo90+m7Wz
# uov96Uugx5wZdewiIIm1ZWTj4lCJHup679LcOa7tAxJMipVaSltQH9fm9TOKczlf
# xtWuBcLU4duZfqwgsILsH7PMkcX1zwQzQD0yAtPhnYz9KNG125bX+iilOe1S8RHq
# v2bbBpMpao4kcUvQI6dMgKRvFmm1eLbhSNOQplDMTGD1tNVdNGkI96jUu+troUjW
# MMi46TQfBAHxtDTpRhIu/87vAVQ8Z6RHhFxesz4Ed5JThaIQRAy6GcO/Jk+QzDzo
# Z0arRIkIsGJ7rZgOVAjx9ctfw8lH9RfjcwB3wdGBYNMNVJqQpUai2Taddf5pXzTZ
# EHIqLEF53SrBjIeInoQrP7U5VlXiMQsxewLdINrAE2l2TR3KBikb+RQRygbTp8jj
# 2yiC0NCUwG+K+ndglN5RMbXjFW6aKa59Xq+b8XzK/DK+AJtgOpHgJv8Qrk62A+tw
# OVLOpjGCA9QwggPQAgEBMHgwYTELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFB1YmxpYyBSU0Eg
# VGltZXN0YW1waW5nIENBIDIwMjACEzMAAABbSrWNQTJt3HQAAAAAAFswDQYJYIZI
# AWUDBAIBBQCgggEtMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG
# 9w0BCQQxIgQgvKowI3Fm4g4qyBATHYmnm277nST25PNgi8ywT+I7XAEwgd0GCyqG
# SIb3DQEJEAIvMYHNMIHKMIHHMIGgBCAvMQNVXZ0b0xxlGw8X/3IEybObuT6a5W1d
# 61CW+cGD7zB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1l
# c3RhbXBpbmcgQ0EgMjAyMAITMwAAAFtKtY1BMm3cdAAAAAAAWzAiBCAAOqoxXdqm
# zgk17tsO7h0hbPloHKhPSeXAKgRnnh+hnjANBgkqhkiG9w0BAQsFAASCAgAr+EiW
# UUkF0S+baitql61Ad4PqGxFqimG5FrNI4bpCniWzfUjbD67Q8jccbi1ZSrUMBMZm
# M5x2j1pIGWpHJ62HzJfvrz4Ft4/8H+EKtXmgt8AyhIQiqE6hMYcOqzLrTDdLS431
# lUx4eZW6rT52c+tZt/nCLENs3HMe7czWIvGUdm91BKmwBEb2GhYLO2cjZODLi88W
# 0pWuPY+1UTySyvCmaM+nojMu7VW/l6dhsma2pRM3UINhDGip5P6z5w7YUZGzR7OF
# F7mqn3OY8lOmq6CZodjHlJ2YT2KYduUO12xutXz390oA1aMmv7HOorsfOtzWq51S
# xXwZSqA48KFwi9u6TcZFewf3yu8x8y/MhGFGXXzK/cSZMC9oMkiJ25Il+OpK/Ej0
# NnrkxKOtXPF8iHiat3qlGbMyPDr8obngmdr9FDosMOe/lJhWptw92+5ZQULFh2CD
# akRwJzOLQEbYaYcUSK6N+jyZLtkuDUhDrpbMoQpnlgoc9cSI/qZ4uGEhrOGOGc7U
# nG+ipYsEMt1U/eDGNisCwkU4KMxpY/re+08zb5NOSaUsRNMhtS1FLCxOnVolRc24
# 3tz7YBmHi/bOfsU9TP0AMLzZpkuYHKHdrf3W/MRAqlqJnJTNaA8OEYkqAV+HyUHz
# DuxLHZHclYMvCXSAVm5iGNQo1sDutpczHO+WgA==
# SIG # End signature block
