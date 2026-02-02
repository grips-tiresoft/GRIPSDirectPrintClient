param(
    [Parameter(Mandatory=$true)][string]$Path
)

# You need the grips-signtool repo installed on your machine in the same level as your this repo
# git clone https://github.com/goodyear/grips-signtool.git
# See GRIPS OneNote for more information

$signToolPath = "$PSScriptRoot\..\grips-signtool\AzureArtifactSigning\microsoft.windows.sdk.buildtools\10.0.26100.7463\bin\10.0.26100.0\x64\signtool.exe"
$DlibDllPath = "$PSScriptRoot\..\grips-signtool\AzureArtifactSigning\microsoft.artifactsigning.client\1.0.115\bin\x64\Azure.CodeSigning.Dlib.dll"
$metaDataFilePath = "$PSScriptRoot\..\grips-signtool\AzureArtifactSigning\GRIPS-codesign.json"

& $signToolPath sign /v /debug /fd SHA256 /tr "http://timestamp.acs.microsoft.com" /td SHA256 /dlib $DlibDllPath /dmdf $metaDataFilePath $Path