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

# Create the file association and file type for .grdp files (system-wide)
. "$ScriptPath\Register-GRDPFileType.ps1"

Write-Host "Installation completed"

Start-Sleep -Seconds 5

Stop-Transcript
# SIG # Begin signature block
# MII7sgYJKoZIhvcNAQcCoII7ozCCO58CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBkRSGOydwaCiIr
# NFUzPxz/2pAxzHUcih1bHGd7B7NBRKCCI9YwggXMMIIDtKADAgECAhBUmNLR1FsZ
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
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQg6cxA
# PELuwPdniUEYc9UCX8ehIfUhcl3uqRj7+SQ48R0wDQYJKoZIhvcNAQEBBQAEggGA
# EkykZLkZ4ZMBywMEnt6fhJZFLL5lzZB90CX7DPwDu4MG3IapzwXvCzNfHPKvQYlY
# JTDkhA7/tBdBJbwEM4WPA260YtwX8Gm4vn2HHXnY/cOr0VmP8IiI118UcAb3sjdq
# E41JcbCfWXRwFp3TPhQ/ReWNicqJXbS0hmQDHo2jO4rq2d8gP9ZUCGBNDI6uxpRZ
# jW6VN+0+BFGcuvDuX3z+iCu8LoS3/oyW8zX2FHQkR5b9UNFlVoSTsPrrAtBRGvRv
# nhyp8PHPQSChZ3QcdyMepapNEyj0ydPpifeiQtSfagOGqJhTuK1DNR4crRrXakyv
# UaP7/+GWGe/Z308GJ2HabLLx/93U6erZzQabvIWtWxXoaj8aFiW6YGoypn6eR97w
# L04m4w+pF9gyXut4p5w/Xw6t7WfuB8MGcBIfbGBNvfFg4nJGa20lP0CYOBm+NAT9
# JirRi3omjsVzRGAHctFMZu2B84AK6AfgCMhxbAllwpmGJHqHt+VHGa94C//I+R9R
# oYIUsjCCFK4GCisGAQQBgjcDAwExghSeMIIUmgYJKoZIhvcNAQcCoIIUizCCFIcC
# AQMxDzANBglghkgBZQMEAgEFADCCAWoGCyqGSIb3DQEJEAEEoIIBWQSCAVUwggFR
# AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIOmUa1D9HJATk6/+f7Ha
# Vuw9D7YWjzBbGqojit+qWn+UAgZpc1aduQYYEzIwMjYwMTI5MTQxODQ3Ljk2MVow
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
# IgQgHJXmkYQ0BhkqIfyhRiW2f50gimOhiKmgTLwZ+MDB5P0wgd0GCyqGSIb3DQEJ
# EAIvMYHNMIHKMIHHMIGgBCD3umpBh3duMPkt6ItkJtOzRwBRt1Q8I513q8lrfqDu
# 3DB8MGWkYzBhMQswCQYDVQQGEwJVUzEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUHVibGljIFJTQSBUaW1lc3RhbXBp
# bmcgQ0EgMjAyMAITMwAAAFxhTACmp/LpIwAAAAAAXDAiBCBk1ueQyCU3fa9rvP45
# Ot7GQN2/MRU0a2lJ5d1+xGgN2DANBgkqhkiG9w0BAQsFAASCAgCieJy/IeTo8fOy
# dMN/79K97+TJIGCBABEKeuNyt1XrycOw7w9C5g5JMW/FXBL0AajInIXqmGfrB2BT
# lpXf+rgotjXQPs0QbfpyQsSJVemzJ1Qc4RRSwB5tgnE50okOnu/xRBvtH1hX+GLq
# ATDkgkrvkUUHcBwDRh90jBw6X9wLlxH2Imi5FxGAO/ln3f2H5BW9/WUaN+MlDlqt
# 66tdm7/Awt81L0qDr4tuiHHCoaxznRTQ2J6A9+ljsI0EDcSJgpMA/DkkOj+umyR6
# lfzSy4SeD+PVH9XpEUW5Jfhnt0yLG2xAbZJYcCk/CdE+STnbBI+T/3ndOQOhpKYK
# MVjg4sLTA4VG+Y/+leX2Z00tUWNsO4N5/kkwln0hTTMsyQfdYNA0t+TJuZrluQZW
# M3+JOt4oFMkyl/6AP/Yat1SWzKXQYp4bfTQeyiG3kHoSIM7llEaU82aoYvZB6rUZ
# 7nUnmtka18zBIdn6eV9csjJvJyTQA4gxpii5FwCEv2XMQgBUZ6wGyQ5nu8jZKgzf
# fXAg4UgabRaDcKyxjd0VsjA1rfNRXSefPHJrLY0Z59ITfoS136Fy7386hGgymJCu
# ZPzLxbbJ9Pq5SjmOGhCmFscsKk++aI/Z/dmf+c+ZrZd4V9xLknAMZFJxeDlSw4yf
# b4xLN7IVJ6LxNUbfpFDw9Ot2eCqggQ==
# SIG # End signature block
