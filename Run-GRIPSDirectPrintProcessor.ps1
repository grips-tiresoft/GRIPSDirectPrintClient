# Version: v1.0.22

# TODO: 
# * DONE: Add RC to GRIPSDirectPrint
#       - Pass to WS when adding updating printers
#       - filter on RC when displaying printers page    
# * DONE: Pick up default printer and use when no printer is specified
# * DONE: Adapt all places where printers are displayed/selected and show normal printers page instead of Web Client Printers when Printing App is disabled
# * DONE: Renumber objects according to GRIPS rules - Viktoras says its OK to use 90000+
# * DONE: Move Settings config.json
# * DONE: Encrypt password using secret protected key read from the registry
# * TODO: Create installation script to install processor as service using nssm
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

# Get decryption key from registry
$key = @(((Get-ItemProperty HKLM:\Software\GRIPS\l02fKiUY).l02fKiUY) -split ",")

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

Get-ScriptVersion -ScriptPath $FullScriptPath

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