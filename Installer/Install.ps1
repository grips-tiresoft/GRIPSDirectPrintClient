$ScriptPath = $PSScriptRoot

Start-Transcript -Path "$ScriptPath\install.log" -Append

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
$config = Get-Content $configFilePath | ConvertFrom-Json

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

$installArgs = "-ExecutionPolicy Bypass -File ""$installPath\Run-GRIPSDirectPrintProcessor.ps1"""
& $nssmPath install "GRIPSDirectPrint Client Service" "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "$installArgs"

# Start the service installed by NSSM
Start-Service -Name "GRIPSDirectPrint Client Service"

# Clean up the extracted temporary directory
Remove-Item -Path $TempExtractPath -Recurse -Force

Write-Host "Installation completed"

Stop-Transcript