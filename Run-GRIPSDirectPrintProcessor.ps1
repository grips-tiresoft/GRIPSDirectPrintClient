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
# MII7sgYJKoZIhvcNAQcCoII7ozCCO58CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDYMGRcSddICfd
# frnidU1Puuw0b2zFMwITGWNYcSo2i6CCI9YwggXMMIIDtKADAgECAhBUmNLR1FsZ
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
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQgL8+E
# vEa3Xe73uGOkZALeDXDFhL74pZxlXJipT6UH5oowDQYJKoZIhvcNAQEBBQAEggGA
# hNTYReYY8/y4n5f3y3ezBAZ7QVDrEvxH5hlXYpmaEWTD8WJ4YdAeCWEyxh4dQaXo
# NbmfFZKRmlNSazmdjriPaF31CCiRDM8wcUxuVAbNxeo2mXpTPFY6wECura/79l2k
# sFqJ4CSwBMQTZt27ThCUIp5o2veFvJp/1E8KTmJ60vbVnkGqsirXR7gd/CZIn8Px
# 6YL0smnuzyHSCO986cSBHCNJPXn2/YJnpwo60q9H/xsNI+LawKDhQ2Ru4ZuzNR0n
# b+t354hSjBsKvnnlGPY4gtQ4MpWW30WQXbrQbXjQgHsDqkbJdrTa5lEm1AyX9ZQj
# qvjBmMMsCrRmwcIp5oxrQDAvahOJ7hzxkfgcfhiYRT33bVTYvnRvbgZyQA24fAT6
# 0xrOcam+wSQr+wMv6PGRmrf+GzbAJfwz6Niq8rWoYLuAgB39ul71BN0q1i3PuYJ2
# iupBV+nJdtIu7OkDAcm4hy5eo+e4bQC9e8ec07QVdnxX8dpUmfMV6Rgqc07WNaw+
# oYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIUizCCFIcC
# AQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSCAVUwggFR
# AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEILquO/r2nmbit0qSH4be
# EywOCegnHiujGkOu/c9XGNHkAgZpc4Ui+9MYEzIwMjYwMTI5MTMyNTE2Ljc2OVow
# BIACAfSggemkgeYwgeMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRl
# ZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjdBMUEtMDVFMC1EOTQ3MTUwMwYD
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
# JYbYbBh7pmgAXVswggefMIIFh6ADAgECAhMzAAAAW0q1jUEybdx0AAAAAABbMA0G
# CSqGSIb3DQEBDAUAMGExCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBQdWJsaWMgUlNBIFRpbWVz
# dGFtcGluZyBDQSAyMDIwMB4XDTI2MDEwODE4NTkwNVoXDTI3MDEwNzE4NTkwNVow
# geMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsT
# JE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMe
# blNoaWVsZCBUU1MgRVNOOjdBMUEtMDVFMC1EOTQ3MTUwMwYDVQQDEyxNaWNyb3Nv
# ZnQgUHVibGljIFJTQSBUaW1lIFN0YW1waW5nIEF1dGhvcml0eTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBAJBUzBbbnlDXee0B0KD5G4/475thFyfctCyu
# ESTWQXvlLi4Wx/td2qUdeq4ideeg6VWhiOHfu3wJV4TUGSRtqh9Ccr1BmiBKv9iu
# FpgHyIBu5Qx38ZsxwlFeXVS+ZqJJKnXRbDNQdcYSoC/6c0hQJ/PH50DBRDQkPXVw
# yFizLrRH9AlrJeUg7BKeT23zftS8/KOJLvEEbHOF6pSOY3ZVprZUWbWjWwRTmoHa
# Q/E8vrWtLNyEJ+b089VW1Ikra3t4GTB5Wby3CL1K2zYnAxBIvafsKMFyj9OuXHcT
# PKMDoFSMeamG9MKOMb6uoG1PjdnDgsLP6EOMRSzrLL7jED1mbB9RSd9fhty+HQr6
# vZgsBn6oUy+YTpNVLskwdtUM82WYAkPztlOt3AiL0qyV7/U3j/uq3vHMjPM0w034
# 0M57Nei0g4BCcMt0dbqoc91VgCb3/36sHQANontn1HOF2oLk8190QRS43isHVra8
# H8sf5+GlqIYsYiCKX04HZiOzZW826nVI6d++8lyTeWmpj90Ua9uPbJhVjwE3oh6t
# O510ySqmSMSLEN07p3Ibe3E6BAb2w93rWzb26+dpSthbKF4kApofqBsWPX4MEtHK
# SOftPmVTCQ47tghrVuHia9jY+Hsj01m4KW4WtkmVm3L6hMZECMa4sjMxAXz+bX/A
# JhWTe6TZAgMBAAGjggHLMIIBxzAdBgNVHQ4EFgQU7/LqUlWWYhXJdXwgYKx4b8Gv
# 0rYwHwYDVR0jBBgwFoAUa2koOjUvSGNAz3vYr0npPtk92yEwbAYDVR0fBGUwYzBh
# oF+gXYZbaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9z
# b2Z0JTIwUHVibGljJTIwUlNBJTIwVGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNy
# bDB5BggrBgEFBQcBAQRtMGswaQYIKwYBBQUHMAKGXWh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwUHVibGljJTIwUlNBJTIw
# VGltZXN0YW1waW5nJTIwQ0ElMjAyMDIwLmNydDAMBgNVHRMBAf8EAjAAMBYGA1Ud
# JQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDBmBgNVHSAEXzBdMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wCAYGZ4EMAQQCMA0GCSqG
# SIb3DQEBDAUAA4ICAQAAH+zd+XKh4OxXYMWFmtgilXAQGctOjCUB1w/uBiC/OXcH
# 3Ia4/XbdUhKzFbaiTbIE6vYZKd1p4u7nKOLkawymAMVyuO7LSl6rLKttZIyLhWjT
# K0zXOz0u4xLq9+bRtBEKJvA6sD5nJwH1IO6z1YizyuIRoalMCnbrUixfWxQn4TAm
# N7t9uk+X2FUThEa3ewzRwhtG+xwaAbLMkxRmR24JnfXd1VxKo90+m7Wzuov96Uug
# x5wZdewiIIm1ZWTj4lCJHup679LcOa7tAxJMipVaSltQH9fm9TOKczlfxtWuBcLU
# 4duZfqwgsILsH7PMkcX1zwQzQD0yAtPhnYz9KNG125bX+iilOe1S8RHqv2bbBpMp
# ao4kcUvQI6dMgKRvFmm1eLbhSNOQplDMTGD1tNVdNGkI96jUu+troUjWMMi46TQf
# BAHxtDTpRhIu/87vAVQ8Z6RHhFxesz4Ed5JThaIQRAy6GcO/Jk+QzDzoZ0arRIkI
# sGJ7rZgOVAjx9ctfw8lH9RfjcwB3wdGBYNMNVJqQpUai2Taddf5pXzTZEHIqLEF5
# 3SrBjIeInoQrP7U5VlXiMQsxewLdINrAE2l2TR3KBikb+RQRygbTp8jj2yiC0NCU
# wG+K+ndglN5RMbXjFW6aKa59Xq+b8XzK/DK+AJtgOpHgJv8Qrk62A+twOVLOpjGC
# A9QwggPQAgEBMHgwYTELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFB1YmxpYyBSU0EgVGltZXN0
# YW1waW5nIENBIDIwMjACEzMAAABbSrWNQTJt3HQAAAAAAFswDQYJYIZIAWUDBAIB
# BQCgggEtMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQgLa5n0aUxG7JKX6Maak/yInEIxHTKp/qG2q1A3Ghc2Q8wgd0GCyqGSIb3DQEJ
# EAIvMYHNMIHKMIHHMIGgBCAvMQNVXZ0b0xxlGw8X/3IEybObuT6a5W1d61CW+cGD
# 7zB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lc3RhbXBp
# bmcgQ0EgMjAyMAITMwAAAFtKtY1BMm3cdAAAAAAAWzAiBCDgfzwUPIIL7e3GO5fa
# 1ORNF5bjYFqbK7wvp/l9d5JL4jANBgkqhkiG9w0BAQsFAASCAgBEKf0Vn48xrzhJ
# dKIxaDDl7jhEkY/p7jpFbdNsn+TNIvhe1AkfJ8gWCivVxMQM171txB8AnLSNCplW
# JrcvnM5gAl1sy7MdrH+dLepIbQLv7WFSccIaeMgcuMoCRuxTVDwQl76FpOZ995Ix
# U2rHNnjLpz+zfVw51FdBXJclgrE/aiFHr6k77/lQgQ4BPf1CKz2T6X49gLWbmuRv
# P0WGeGQcwbFKZaMvaR1ZVp1PNurPDYpklYvYLlQe+/wpB8dz+y9kBy2rZHWra0+C
# Wi2zVlyC0srqA7iWe66K0Zap7Rzup+9vcSExQNbVXQ7Azd0Ct0Iyz9GWy3EDDaLG
# GBmQxKyeskXqGFV05zjTPu8SVMRD4JIIBNnba1RS4zFAamifBKZm9ERrL6i694u5
# APzfxNJ75PG62E+vZ/bkiv5D6gBGENPnzzdEs9zNE26km4m5mkB2GIFIHYP/l8A6
# vLzID91PdO4jnpL5SHJdWRSP16mBS7WAjULEih3MFv6KIj9yQw1hRkwcd+x5A9Bm
# p4MZFAX3uMo+HgfAZkTh1dg2/vrjEFb8Yd8zgR8cXFa0ezX5i8J0ZxFYWAGk1Q2F
# H8l0gCmBblmVSnMH/rkpL7LH8EHhcpKpNqRIbLoVxpOpMl5id0CVcClycBFDCSA9
# uTpsQUrq2hwbr9fytOk6V5z6jKi2Sw==
# SIG # End signature block
