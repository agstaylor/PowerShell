# Checks for possibly misplaced or duplicated certs
#
# *Check Trusted Root Store
# First get all trusted root certs in store
$rootcerts = Get-Childitem 'cert:\LocalMachine\root' -Recurse
# Check for certs which were not their own issuer (thus not a root certificate!)
$misplacedrootcerts = $rootcerts | Where-Object {$_.Issuer -ne $_.Subject}
Foreach ($cert in $misplacedrootcerts) {
    if (($intermediatecerts).thumbprint -contains $cert.thumbprint) {
        Write-Host -ForegroundColor:Yellow "Intermediate cert found duplicated in root cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Delete certificate from trusted root store."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
    else {
        Write-Host -ForegroundColor:Yellow "Intermediate cert found in root cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Move certificate from trusted root store to intermediate store."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
}
 
# *Check Trusted Intermediate Store
# First get all trusted intermediate certs in store
$intermediatecerts = Get-Childitem 'cert:\LocalMachine\CA' -Recurse | Where {$_.Subject -ne 'CN=Root Agency'}
# Check for certs which issued themselves (thus not an intermediate certificate!)
$misplacedintermediatecerts = $intermediatecerts | Where-Object {$_.Issuer -eq $_.Subject}
Foreach ($cert in $misplacedintermediatecerts) {
    if (($rootcerts).thumbprint -contains $cert.thumbprint) {
        Write-Host -ForegroundColor:Yellow "Trusted root cert found duplicated in intermediate cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Delete certificate from intermediate store."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
    else {
        Write-Host -ForegroundColor:Yellow "Trusted root cert found in intermediate cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Move certificate from intermediate cert store to root cert store."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
}
 
# *Check Local machine store
# First get all local machine certs in store
$mycerts = Get-Childitem 'cert:\LocalMachine\My' -Recurse
$myselfsignedcerts = $mycerts | Where-Object {$_.Issuer -eq $_.Subject}
$myrootduplicatedcerts = $mycerts | Where-Object {($rootcerts).thumbprint -contains $_.thumbprint}
$myintermediateduplicatedcerts = $mycerts | Where-Object {($intermediatecerts).thumbprint -contains $_.thumbprint}
 
Foreach ($cert in $myrootduplicatedcerts) {
    if (($myselfsignedcerts).thumbprint -contains $cert.thumbprint) {
        Write-Host -ForegroundColor:Yellow "Local machine certificate found duplicated in trusted root cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Yellow "Certificate status: SELF-SIGNED (Possible trusted root authority)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Validate if the cert is a trusted root or simply self-signed and remove duplicate from one of the stores."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
    else {
        Write-Host -ForegroundColor:Yellow "Local machine certificate found duplicated in trusted root cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Yellow "Certificate status: NOT SELF-SIGNED (Not a possible trusted root authority)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Delete this certificate from the trusted root store."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
}
 
Foreach ($cert in $myintermediateduplicatedcerts) {
    if (($myselfsignedcerts).thumbprint -contains $cert.thumbprint) {
        Write-Host -ForegroundColor:Yellow "Local machine certificate found duplicated in intermediate root cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Yellow "Certificate status: SELF-SIGNED (Possible trusted root authority)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Delete this certificate from the intermediate root store. Possibly move it to the trusted root store."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
    else {
        Write-Host -ForegroundColor:Yellow "Local machine certificate found duplicated in trusted root cert store - $($cert.Subject)"
        Write-Host -ForegroundColor:Yellow "Certificate status: NOT SELF-SIGNED (Not a possible trusted root authority)"
        Write-Host -ForegroundColor:Magenta "Recommended action: Validate if the cert is an intermediate root or a valid local machine cert and remove the duplicate from one of the stores."
        Write-Host -ForegroundColor:DarkMagenta '**Certificate Details **'
        Write-Host $cert
        Read-Host -Prompt ''
    }
}