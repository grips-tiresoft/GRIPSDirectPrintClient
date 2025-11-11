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
        1
        # Validate selection
        $selInt = $selection -as [int]
        $isValid = ($selInt -ge $StartIndex) -and ($selInt -le $($StartIndex + $Items.Count - 1))

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
            if ($Method -ne "Post") {
                $URL = "$URL/Company('$($Authentication.Company)')"
            }
        }

        $URL = "$URL/$WebServiceName"

        if (-not ([string]::IsNullOrEmpty($Authentication.Company))) {
            if ($Method -eq "Post") {
                $URL = "$($URL)?company='$($Authentication.Company)'"
            }
        }

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

function Copy-ScriptFolder {
    # Copy the extracted files from the sub-folder to the destination directory
    $resolvedPath = Resolve-Path -Path "$ScriptPath\..\"

    # Ensure the destination directory exists
    if (-Not (Test-Path -Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath
    }

    # Copy the contents of the source directory to the destination directory
    Copy-Item -Path "$resolvedPath\*" -Destination $installPath -Recurse -Force

    # Set ACL to allow Users full control
    # Use SID instead of group name to support all Windows language versions
    $usersSid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-545")
    $acl = Get-Acl -Path $installPath
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($usersSid, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    Set-Acl -Path $installPath -AclObject $acl
}


#TODO: Load localized strings from a resource file
#Import-LocalizedData -BindingVariable strings -FileName GRIPSDirectPrint-InstallStrings.psd1 -BaseDirectory ".\Resources"

function Install-GRIPSDirectPrintClientService {
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
    }
    else {
        $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title
    }

    # Convert the selection to the country code
    $selectedDatabase = @($selectedCountry.Databases.PSObject.Properties.Name)[$selectionResult.Index]
    $selectedBaseURL = $selectedCountry.Databases.$selectedDatabase.BaseURL
    $selectedCompany = $selectedCountry.Databases.$selectedDatabase.Company

    Write-Host "Selected Database: $($selectedDatabase) ($selectedBaseURL)"

    $keyPath = "$ScriptPath\l02fKiUY\l02fKiUY.txt"

    $key = @(((Get-Content $keyPath) -split ","))

    $credFile = "$installPath\$($jsonContent.BasicAuthLogin).TXT"

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

    $GetCompaniesWS = "GRIPSDirectPrintGeneralWS_GetCompanies"

    Clear-Host  # Clears the console

    # Ask user for the UserName that will be used to filter the companies
    do {
        Clear-Host  # Clears the console

        $UserName = Read-Host "Please enter your UserName (to filter the list of companies)"
        $isValid = -not [string]::IsNullOrEmpty($UserName)
        if (-not $isValid) {
            Write-Host "Invalid entry. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2  # Give user time to read the message before clearing
        }
    } while (-not $isValid)

    # Fetch the list of companies
    $Body = "{""userName"": $($UserName | ConvertTo-Json) }"

    $Companies = (Invoke-BCWebService -Method Post -BaseURL $selectedBaseURL -WebServiceName $GetCompaniesWS -Authentication $Authentication -Body $Body).value
    $CompaniesObject = $Companies | ConvertFrom-Json

    # Present the list of companies to the user and ask for a selection
    $title = "Company Selection"

    $Items = [ordered]@{}
    foreach ($Company in $CompaniesObject.companies) {
        $Items.Add($Company.Companyname, $Company.DisplayName)
    }

    $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title
    $selectedCompany = $selectionResult.Key
    Write-Host "Selected Company: $($selectedCompany)"

    # Set Authentication using selected company
    $Authentication = @{
        #"Company"                     = 'NAS Company' # Note: Must exist or be left empty if a Default Company is setup in the Service Tier. Only used for authentication as printers and jobs are PerCompany=false
        "Company"                     = $selectedCompany

        "BasicAuthLogin"              = $jsonContent.BasicAuthLogin;
        "BasicAuthPassword"           = $(([Net.NetworkCredential]::new('', $credential.Password).Password))

        "OAuth2CustomerAADIDOrDomain" = $jsonContent.OAuth2CustomerAADIDOrDomain
        "OAuth2ClientID"              = $jsonContent.OAuth2ClientID
        "OAuth2ClientSecret"          = $jsonContent.OAuth2ClientSecret
    }

    $GetRespCentersWS = "GRIPSDirectPrintGeneralWS_GetResponsibilityCenters"


    # Fetch the list of responsibility centers
    $RespCenters = (Invoke-BCWebService -Method Post -BaseURL $selectedBaseURL -WebServiceName $GetRespCentersWS -Authentication $Authentication).value
    $RespCentersObject = $RespCenters | ConvertFrom-Json

    # Present the list of resonsibility centers to the user and ask for a selection
    $Items = [ordered]@{}
    $title = "Responsibility Center Selection"
    foreach ($RespCtr in $RespCentersObject.responsibilityCenters) {
        $Items.Add($RespCtr.RespCenterCode, $RespCtr.RespCenterName)
    }

    $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title
    $selectedRespCtr = $selectionResult.Key
    Write-Host "Selected Responsibility Center: $($selectedRespCtr)"
    # Write Company and BaseURL to userconfig.json
    $userConfigPath = "$installPath\userconfig.json"
    $userConfig = @{
        Company = $selectedCompany
        BaseURL = $selectedBaseURL 
        RespCtr = $selectedRespCtr
        UsePrereleaseVersion = $false
    } | ConvertTo-Json -Depth 4

    $userConfig | Out-File -FilePath $userConfigPath -Encoding UTF8

    # Define the path to the NSSM executable
    $nssmPath = "$installPath\Installer\nssm-2.24\win64\nssm.exe"

    # Install the client service using NSSM
    $installArgs = "-ExecutionPolicy Bypass -File ""$installPath\Run-GRIPSDirectPrintProcessor.ps1"""
    & $nssmPath install "GRIPSDirectPrint Client Service" "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "$installArgs"

    # Start the service installed by NSSM
    Start-Service -Name "GRIPSDirectPrint Client Service"
}    

$user = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
#$isAdmin = $true

if (-Not $isAdmin) {
    $NotAdminError = "Script is not running with administrative privileges..attempting to relaunch elevated"
    Write-Output -ForegroundColor Red $NotAdminError
    Start-Sleep -s 2
    #Write-Error -Message $NotAdminError -ErrorAction Stop
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}
else {
    $IsAdminMsg = "Script is running with administrative privileges - Installing GRIPSDirectPrint Client..."
    Write-Output $IsAdminMsg
}

if (-Not $isAdmin) {
    $NotAdminError = "Script is not running with administrative privileges..GRIPSDirectPrint Client is not installed"
    Write-Output -ForegroundColor Red $NotAdminError
    Start-Sleep -s 2
    Write-Error -Message $NotAdminError -ErrorAction Stop
}

$ScriptPath = $PSScriptRoot

Start-Transcript -Path "$ScriptPath\install.log" -Append

Write-Host "Starting GRIPSDirectPrint Client installation..." -ForegroundColor White
Write-Host  

# Set PowerShell console to use UTF-8 encoding
[Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

# Define the path to the configuration file
$configFilePath = "$ScriptPath\install.json"

# Check if the configuration file exists
if (-Not (Test-Path -Path $configFilePath)) {
    Write-Error "Configuration file not found at path: $configFilePath"
    exit
}

# Load configuration from JSON file
$config = Get-Content $configFilePath -Encoding UTF8 | ConvertFrom-Json

#$releaseApiUrl = $config.ReleaseApiUrl;
$installPath = $config.InstallPath;

#Write-Host "Downloading client..." -ForegroundColor White

#$LatestRelease = Invoke-RestMethod -Uri $releaseApiUrl -Method Get
    
#$TempZipFile = [System.IO.Path]::GetTempFileName() + ".zip"
#$TempExtractPath = [System.IO.Path]::GetTempPath() + [System.Guid]::NewGuid().ToString()

# Get the URL of the source code zip
#$downloadUrl = $LatestRelease.zipball_url

# Download the ZIP file containing the new script version and other files
#Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZipFile

# Extract the ZIP file to a temporary directory
#Expand-Archive -Path $TempZipFile -DestinationPath $TempExtractPath

# Find the sub-folder in the extracted directory
#$extractedSubFolder = Get-ChildItem -Path $TempExtractPath | Where-Object { $_.PSIsContainer } | Select-Object -First 1

# Clean up temporary files
#Remove-Item -Path $TempZipFile -Force

Copy-ScriptFolder

Stop-Transcript

$ScriptPath = "$installPath\Installer"

Start-Transcript -Path "$ScriptPath\install.log" -Append

$installConfigPath = "$ScriptPath\install.json"

# Ask user whether to install the GRIPSDirectPrintClient service
try {
    $choices = @(
        (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Install the GRIPSDirectPrintClient service."),
        (New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not install the service.")
    )
    $decision = $Host.UI.PromptForChoice("GRIPSDirectPrintClient Service", "Do you want to install the GRIPSDirectPrintClient service?", $choices, 1)
} catch {
    # Fallback for non-interactive hosts
    $answer = Read-Host "Do you want to install the GRIPSDirectPrintClient service? (Y/N)"
    $decision = if ($answer -match '^(?i)y(?:es)?$') { 0 } else { 1 }
}

if ($decision -eq 0) {
    if (Get-Command -Name Install-GRIPSDirectPrintClientService -ErrorAction SilentlyContinue) {
        Write-Host "Installing GRIPSDirectPrintClient service..."
        try {
            Install-GRIPSDirectPrintClientService
            Write-Host "GRIPSDirectPrintClient service installation completed."
        } catch {
            Write-Error "Failed to install GRIPSDirectPrintClient service: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Install-GRIPSDirectPrintClientService command not found. Skipping service installation."
    }
} else {
    Write-Host "Skipping GRIPSDirectPrintClient service installation."
    $userConfigPath = "$installPath\userconfig.json"
    $userConfig = @{
        UsePrereleaseVersion = $false
    } | ConvertTo-Json -Depth 4

    $userConfig | Out-File -FilePath $userConfigPath -Encoding UTF8
}

# Now using .sig files which are automatically associated with Signotec SignoSign2
# Create the file association for .signpdf files
#Write-Host "cmd.exe /c ""assoc $($config.Sign_ext)=SignedPDFFile"""
#& cmd.exe /c "assoc $($config.Sign_ext)=SignedPDFFile"

#Write-Host "cmd.exe /c ftype SignedPDFFile=""""$($config.Sign_exe)"""" """"$($config.Sign_params)"""""
#& cmd.exe /c ftype SignedPDFFile="""$($Config.Sign_exe)""" ""$config.Sign_params""

# Create the file association and file type for .grdp files
Write-Host 'cmd.exe /c "assoc .grdp=GRIPS.DirectPrint.Archive"'
& cmd.exe /c "assoc .grdp=GRIPS.DirectPrint.Archive"

Write-Host "ftype GRIPS.DirectPrint.Archive=wscript.exe ""$installPath\Print-GRDPFile.vbs"" ""%1"""
& cmd.exe /c "ftype GRIPS.DirectPrint.Archive=wscript.exe ""$installPath\Print-GRDPFile.vbs"" ""%1"""

# Clean up the extracted temporary directory
#Remove-Item -Path $TempExtractPath -Recurse -Force

Write-Host "Installation completed"

Start-Sleep -Seconds 5

Stop-Transcript
# SIG # Begin signature block
# MIIP2AYJKoZIhvcNAQcCoIIPyTCCD8UCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqwze5CWhIMadMxQJ8Zc0DqMW
# 3dCggg0oMIIFUzCCBDugAwIBAgITGAAAFn2/inOgnDqjKgAAAAAWfTANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUAWtQ+U8ncROc8oEjvlXI77/rv2EwDQYJKoZI
# hvcNAQEBBQAEggEA0nGOaX8ncyD3JrVkt0zJu79Pv5zi7eo4dNhpJo6qNvd5cP9H
# DbnRrn9qHxDF57PTRl/8ittIR8Z+bTpVLsLQ3vy6M8M+GnLJQpBCnX2yrG03f8Nk
# AqwykhiuQWdOwyqlcmdwUijZx7y82hedgFEzwxYDip8bui33xAhIQm00s36vsbz0
# lHHkH5BZ1t5CmNWW2uP9iPU2HsHXY2lDWOWzahgK4rBcv2ChGB8VOU4b97H+4iJR
# 8gJrFqkRVm2Ad3i/+GegRjKtLl8MyrrG/VLVA7dYfqJptqHYTz/n6cJavRHiawhf
# RS57mr5fhV9qtnPiUtXqtIgaSgVqbdaRyfYCOQ==
# SIG # End signature block
