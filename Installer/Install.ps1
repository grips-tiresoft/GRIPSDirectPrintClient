$ScriptPath = $PSScriptRoot

Start-Transcript -Path "$ScriptPath\install.log" -Append

function DisplayMenuAndReadSelection {
    param (
        [System.Collections.Specialized.OrderedDictionary]$Items,
        [string]$Title,
        [int]$StartIndex = 1
    )

    do {
        Clear-Host  # Clears the console

        # Display the title if provided
        if (-not [string]::IsNullOrWhiteSpace($Title)) {
            Write-Host $Title -ForegroundColor Cyan
            Write-Host ('=' * $Title.Length) -ForegroundColor Cyan  # Optional: underline for the title
        }

        # Display menu options
        $index = $StartIndex
        foreach ($key in $Items.Keys) {
            Write-Host "$index`t$key`t$($Items[$key])"
            $index++
        }

        # Prompt for user input
        $selection = Read-Host "Please select an option ($($StartIndex)-$($StartIndex+$Items.Count-1))"

        # Validate selection
        $selInt = $selection -as [int]
        $isValid = ($selInt -ge $StartIndex) -and ($selInt -le $($StartIndex+$Items.Count-1))

        if (-not $isValid) {
            Write-Host "Invalid selection. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2  # Give user time to read the message before clearing
        }

    } while (-not $isValid)

    # Convert selection to corresponding dictionary key
    $selectedKey = ($Items.Keys | Select-Object -Index ($selection - $StartIndex))
    Write-Host "You selected: $selectedKey" -ForegroundColor Green

    # Convert selection to corresponding dictionary key and index
    $selectedKey = ($Items.Keys | Select-Object -Index ($selection - $StartIndex))
    $selectedIndex = [int]$selection - $StartIndex  # Adjust for zero-based index if needed

    # Return both selected key and index
    return @{ "Key" = $selectedKey; "Index" = $selectedIndex }
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

# Set PowerShell console to use UTF-8 encoding
[Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

# Define the path to the configuration file
$configFilePath = "$ScriptPath\install.json"

# Define the path to the NSSM executable
$nssmPath = "$ScriptPath\nssm-2.24\win64\nssm.exe"

# Check if the configuration file exists
if (-Not (Test-Path -Path $configFilePath)) {
    Write-Error "Configuration file not found at path: $configFilePath"
    exit
}

# Load configuration from JSON file
$config = Get-Content $configFilePath -Encoding UTF8 | ConvertFrom-Json

$releaseApiUrl = $config.ReleaseApiUrl;
$installPath = $config.InstallPath;

Write-Host "Downloading client..." -ForegroundColor White

$LatestRelease = Invoke-RestMethod -Uri $releaseApiUrl -Method Get
    
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

# Clean up temporary files
Remove-Item -Path $TempZipFile -Force

# Copy the extracted files from the sub-folder to the destination directory
$resolvedPath = Resolve-Path -Path $extractedSubFolder.FullName

# Ensure the destination directory exists
if (-Not (Test-Path -Path $installPath)) {
    New-Item -ItemType Directory -Path $installPath
}

# Copy the contents of the source directory to the destination directory
Copy-Item -Path "$resolvedPath\*" -Destination $installPath -Recurse -Force

# Set ACL to allow Users full control
$acl = Get-Acl -Path $installPath
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
$acl.SetAccessRule($accessRule)
Set-Acl -Path $installPath -AclObject $acl

# TODO: 
$installConfigPath = "$resolvedPath\Installer\install.json"

# Load configuration from JSON file
$jsonContent = Get-Content $installConfigPath -Encoding UTF8 | ConvertFrom-Json

# Extract countries and prepare the options for PromptForChoice
$Items = [ordered]@{}

foreach ($countryCode in $jsonContent.Countries.PSObject.Properties.Name) {
    $country = $jsonContent.Countries.$countryCode
    $Items.Add("$countryCode", "$($country.Name)")
}

# Present the countries to the user and ask for a selection
$title = "Country Selection"
#$selectedOption = $Host.UI.PromptForChoice($title, $message, $options, 0)
$selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title

# Convert the selection to the country code
$selectedCountryCode = $jsonContent.Countries.PSObject.Properties.Name[$selectionResult.Index]
$selectedCountry = $jsonContent.Countries.$selectedCountryCode

Write-Host "Selected Country: $($selectedCountry.Name) ($selectedCountryCode)"

# Extract database and prepare the options for PromptForChoice
$Items = [ordered]@{}

foreach ($DatabaseName in $selectedCountry.Databases.PSObject.Properties.Name) {
    $label = "$DatabaseName"
    $helpMessage = ""
    $Items.Add($label, $helpMessage)
}

# Present the database to the user and ask for a selection
$title = "Database Selection"
#$selectedOption = $Host.UI.PromptForChoice($title, $message, $options, 0)
if ($selectedCountry.StartIndex) {
    $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title -StartIndex $selectedCountry.StartIndex
} else {
    $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title
}

# Convert the selection to the country code
$selectedDatabase = $selectedCountry.Databases.PSObject.Properties.Name[$selectionResult.Index]
$selectedBaseURL = $selectedCountry.Databases.$selectedDatabase.BaseURL
$selectedCompany = $selectedCountry.Databases.$selectedDatabase.Company

Write-Host "Selected Database: $($selectedDatabase) ($selectedBaseURL)"

# Get decryption key from registry
$key = @(((Get-ItemProperty HKLM:\Software\GRIPS\l02fKiUY).l02fKiUY) -split ",")

$credFile = "$resolvedPath\$($jsonContent.BasicAuthLogin).TXT"

$credential = Get-StoredCredential -credFile $credFile -key $key

# Authentication:
$Authentication = @{
    #"Company"                     = 'NAS Company' # Note: Must exist or be left empty if a Default Company is setup in the Service Tier. Only used for authentication as printers and jobs are PerCompany=false
    "Company"                     = $selectedCompany

    "BasicAuthLogin"              = $jsonContent.BasicAuthLogin;
    "BasicAuthPassword"           = $(([Net.NetworkCredential]::new('', $credential.Password).Password))

    "OAuth2CustomerAADIDOrDomain" = $jsonContent.OAuth2CustomerAADIDOrDomain
    "OAuth2ClientID"              = $jsonContent.OAuth2ClientID
    "OAuth2ClientSecret"          = $jsonContent.OAuth2ClientSecret
}

$CompaniesWS = "Companies"

# Fetch the list of companies
$Companies = (Invoke-BCWebService -Method Get -BaseURL $selectedBaseURL -WebServiceName $CompaniesWS -Authentication $Authentication).value

Return(1)

# Write Company and BaseURL to userconfig.json
$userConfigPath = "$installPath\userconfig.json"
$userConfig = @{
    Company = $companies | Out-GridView -Title "Select Company" -PassThru
    BaseURL = $config.BaseUrl -f $selectedCountry, $selectedDatabase
} | ConvertTo-Json -Depth 4
$userConfig | Out-File -FilePath $userConfigPath -Encoding UTF8

# Install the client service using NSSM
$installArgs = "-ExecutionPolicy Bypass -File ""$installPath\Run-GRIPSDirectPrintProcessor.ps1"""
& $nssmPath install "GRIPSDirectPrint Client Service" "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "$installArgs"

# Start the service installed by NSSM
Start-Service -Name "GRIPSDirectPrint Client Service"

# Clean up the extracted temporary directory
Remove-Item -Path $TempExtractPath -Recurse -Force

Write-Host "Installation completed"

Stop-Transcript