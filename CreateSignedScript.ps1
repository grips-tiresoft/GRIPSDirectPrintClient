param(
    [Parameter(Mandatory=$true)][string]$Path
)

# You need the certificate installed on your machine to sign the script
# See GRIPS OneNote for more information

$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Subject -like "*GRIPS_Codesign.goodyear.com*"}
Set-AuthenticodeSignature -FilePath $Path -Certificate $cert
