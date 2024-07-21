param (
    [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
    [String]$File,
    [string]$configFile = "$PSScriptRoot\config.json",
    [string]$userConfigFile = "$PSScriptRoot\userconfig.json"
)

Write-Host "Signing file: $File"

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

#$Sign_exe = "C:\Program Files\signotec\signoSign2\SignoSign2.exe"
#$Sign_params = """{0}""" -f $File

$Sign_exe = $config.Sign_exe
$Sign_params = $config.Sign_params -f $File

Start-Process -FilePath $Sign_exe -ArgumentList "$Sign_params" -Wait -NoNewWindow

#Read-Host -Prompt "Press Enter to continue"