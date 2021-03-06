#region assembly import
Add-Type -Path $PSScriptRoot\Library\SysadminsLV.Asn1Parser.dll -ErrorAction Stop
Add-Type -Path $PSScriptRoot\Library\PKI.Core.dll -ErrorAction Stop
Add-Type -AssemblyName System.Security -ErrorAction Stop
#endregion

#region helper functions
function __RestartCA ($ComputerName) {
	$wmi = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "name='certsvc'"
	if ($wmi.State -eq "Running") {
		[void]$wmi.StopService()
		while ((Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "name='CertSvc'" -Property "State").State -ne "Stopped") {
			Write-Verbose "Waiting for 'CertSvc' service stop."
			Start-Sleep 1
		}
		[void]$wmi.StartService()
	}
}

function Test-XCEPCompat {
	if (
		[Environment]::OSVersion.Version.Major -lt 6 -or
		([Environment]::OSVersion.Version.Major -eq 6 -and
		[Environment]::OSVersion.Version.Minor -lt 1)
	) {$false} else {$true}
}

function Ping-Wmi ($ComputerName) {
	$success = $true
	try {[wmiclass]"\\$ComputerName\root\DEFAULT:StdRegProv"}
	catch {$success = $false}
	$success
}

function Ping-ICertAdmin ($ConfigString) {
	$success = $true
	[void]($ConfigString -match "(.+)\\(.+)")
	$hostname = $matches[1]
	$caname = $matches[2]
	try {
		$CertAdmin = New-Object -ComObject CertificateAuthority.Admin
		$var = $CertAdmin.GetCAProperty($ConfigString,0x6,0,4,0)
	} catch {$success = $false}
	$success
}

function Write-ErrorMessage {
	param (
		[PKI.Utils.PSErrorSourceEnum]$Source,
		$ComputerName,
		$ExtendedInformation
	)
$DCUnavailable = @"
"Active Directory domain could not be contacted.
"@
$CAPIUnavailable = @"
Unable to locate required assemblies. This can be caused if attempted to run this module on a client machine where AdminPack/RSAT (Remote Server Administration Tools) are not installed.
"@
$WmiUnavailable = @"
Unable to connect to CA server '$ComputerName'. Make sure if Remote Registry service is running and you have appropriate permissions to access it.
Also this error may indicate that Windows Remote Management protocol exception is not enabled in firewall.
"@
$XchgUnavailable = @"
Unable to retrieve any 'CA Exchange' certificates from '$ComputerName'. This error may indicate that target CA server do not support key archival. All requests which require key archival will immediately fail.
"@
	switch ($source) {
		DCUnavailable {
			Write-Error -Category ObjectNotFound -ErrorId "ObjectNotFoundException" `
			-Message $DCUnavailable
		}
		CAPIUnavailable {
			Write-Error -Category NotImplemented -ErrorId "NotImplementedException" `
			-Message $NoCAPI; exit
		}
		CAUnavailable {
			Write-Error -Category ResourceUnavailable -ErrorId ResourceUnavailableException `
			-Message "Certificate Services are either stopped or unavailable on '$ComputerName'."
		}
		WmiUnavailable {
			Write-Error -Category ResourceUnavailable -ErrorId ResourceUnavailableException `
			-Message $WmiUnavailable
		}
		WmiWriteError {
			try {$text = Get-ErrorMessage $ExtendedInformation}
			catch {$text = "Unknown error '$code'"}
			Write-Error -Category NotSpecified -ErrorId NotSpecifiedException `
			-Message "An error occured during CA configuration update: $text"
		}
		ADKRAUnavailable {
			Write-Error -Category ObjectNotFound -ErrorId "ObjectNotFoundException" `
			-Message "No KRA certificates found in Active Directory."
		}
		ICertAdminUnavailable {
			Write-Error -Category ResourceUnavailable -ErrorId ResourceUnavailableException `
			-Message "Unable to connect to management interfaces on '$ComputerName'"
		}
		NoXchg {
			Write-Error -Category ObjectNotFound -ErrorId ObjectNotFoundException `
			-Message $XchgUnavailable
		}
		NonEnterprise {
			Write-Error -Category NotImplemented -ErrorAction NotImplementedException `
			-Message "Specified Certification Authority type is not supported. The CA type must be either 'Enterprise Root CA' or 'Enterprise Standalone CA'."
		}
	}
}
#endregion

#region module-scope variable definition
# define Configuration naming context DN path
#$ConfigContext = ([ADSI]"LDAP://RootDSE").ConfigurationNamingContext
try {
	$Domain = "CN=Configuration,DC=" + ([DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Forest.Name -replace "\.",",DC=")
	$ConfigContext = "CN=Public Key Services,CN=Services," + $Domain
	$NoDomain = $false
} catch {$NoDomain = $true}
$RegPath = "System\CurrentControlSet\Services\CertSvc\Configuration"
# check whether ICertAdmin CryptoAPI interfaces are available. The check is not performed when
# only client part is installed.
if (Test-Path $PSScriptRoot\Server) {
	try {$CertAdmin = New-Object -ComObject CertificateAuthority.Admin}
	catch {Write-ErrorMessage -Source "CAPIUnavailable"}
}
$Win2003	= if ([Environment]::OSVersion.Version.Major -lt 6) {$true} else {$false}
$Win2008	= if ([Environment]::OSVersion.Version.Major -eq 6 -and [Environment]::OSVersion.Version.Minor -eq 0) {$true} else {$false}
$Win2008R2	= if ([Environment]::OSVersion.Version.Major -eq 6 -and [Environment]::OSVersion.Version.Minor -eq 1) {$true} else {$false}
$Win2012	= if ([Environment]::OSVersion.Version.Major -eq 6 -and [Environment]::OSVersion.Version.Minor -eq 2) {$true} else {$false}
$Win2012R2	= if ([Environment]::OSVersion.Version.Major -eq 6 -and [Environment]::OSVersion.Version.Minor -eq 3) {$true} else {$false}

$RestartRequired = @"
New {0} are set, but will not be applied until Certification Authority service is restarted.
In future consider to use '-RestartCA' switch for this cmdlet to restart Certification Authority service immediatelly when new settings are set.

See more: Start-CertificationAuthority, Stop-CertificationAuthority and Restart-CertificationAuthority cmdlets.
"@
$NothingIsSet = @"
Input object was not modified since it was created. Nothing is written to the CA configuration.
"@
#endregion

#region module installation stuff
# dot-source all function files
Get-ChildItem -Path $PSScriptRoot -Include *.ps1 -Recurse | Foreach-Object { . $_.FullName }
$aliases = @()
if ($Win2008R2 -and (Test-Path $PSScriptRoot\Server)) {
	New-Alias -Name Add-CEP					-Value Add-CertificateEnrollmentPolicyService -Force
	New-Alias -Name Add-CES					-Value Add-CertificateEnrollmentService -Force
	New-Alias -Name Remove-CEP				-Value Remove-CertificateEnrollmentPolicyService -Force
	New-Alias -Name Remove-CES				-Value Remove-CertificateEnrollmentService -Force
	$aliases += "Add-CEP", "Add-CES", "Remove-CEP", "Remove-CES"
}
if (($Win2008 -or $Win2008R2) -and (Test-Path $PSScriptRoot\Server)) {
	New-Alias -Name Install-CA				-Value Install-CertificationAuthority -Force
	New-Alias -Name Uninstall-CA			-Value Uninstall-CertificationAuthority -Force
	$aliases += "Install-CA", "Uninstall-CA"
}

if (!$NoDomain -and (Test-Path $PSScriptRoot\Server)) {
	New-Alias -Name Get-CA					-Value Get-CertificationAuthority -Force
	New-Alias -Name Get-KRAFlag				-Value Get-KeyRecoveryAgentFlag -Force
	New-Alias -Name Enable-KRAFlag			-Value Enable-KeyRecoveryAgentFlag -Force
	New-Alias -Name Disable-KRAFlag			-Value Disable-KeyRecoveryAgentFlag -Force
	New-Alias -Name Restore-KRAFlagDefault	-Value Restore-KeyRecoveryAgentFlagDefault -Force
	$aliases += "Get-CA", "Get-KRAFlag", "Enable-KRAFlag", "Disable-KRAFlag", "Restore-KRAFlagDefault"
}
if (Test-Path $PSScriptRoot\Server) {
	New-Alias -Name Connect-CA					-Value Connect-CertificationAuthority -Force
	
	New-Alias -Name Add-AIA						-Value Add-AuthorityInformationAccess -Force
	New-Alias -Name Get-AIA						-Value Get-AuthorityInformationAccess -Force
	New-Alias -Name Remove-AIA					-Value Remove-AuthorityInformationAccess -Force
	New-Alias -Name Set-AIA						-Value Set-AuthorityInformationAccess -Force

	New-Alias -Name Add-CDP						-Value Add-CRLDistributionPoint -Force
	New-Alias -Name Get-CDP						-Value Get-CRLDistributionPoint -Force
	New-Alias -Name Remove-CDP					-Value Remove-CRLDistributionPoint -Force
	New-Alias -Name Set-CDP						-Value Set-CRLDistributionPoint -Force
	
	New-Alias -Name Get-CRLFlag					-Value Get-CertificateRevocationListFlag -Force
	New-Alias -Name Enable-CRLFlag				-Value Enable-CertificateRevocationListFlag -Force
	New-Alias -Name Disable-CRLFlag				-Value Disable-CertificateRevocationListFlag -Force
	New-Alias -Name Restore-CRLFlagDefault		-Value Restore-CertificateRevocationListFlagDefault -Force
	
	New-Alias -Name Remove-Request				-Value Remove-DatabaseRow -Force
	
	New-Alias -Name Get-CAACL					-Value Get-CASecurityDescriptor -Force
	New-Alias -Name Add-CAACL					-Value Add-CAAccessControlEntry -Force
	New-Alias -Name Remove-CAACL				-Value Remove-CAAccessControlEntry -Force
	New-Alias -Name Set-CAACL					-Value Set-CASecurityDescriptor -Force
	$aliases += "Connect-CA", "Add-AIA", "Get-AIA", "Remove-AIA", "Add-CDP", "Get-CDP", "Remove-CDP",
		"Set-CDP", "Get-CRLFlag", "Enable-CRLFlag", "Disable-CRLFlag", "Restore-CRLFlagDefault",
		"Remove-Request", "Get-CAACL", "Add-CAACL", "Remove-CAACL", "Set-CAACL"
}

if (Test-Path $PSScriptRoot\Client) {
	New-Alias -Name "oid"						-Value Get-ObjectIdentifier -Force
	New-Alias -Name oid2						-Value Get-ObjectIdentifierEx -Force

	New-Alias -Name Get-Csp						-Value Get-CryptographicServiceProvider -Force

	New-Alias -Name Get-CRL						-Value Get-CertificateRevocationList -Force
	New-Alias -Name Show-CRL					-Value Show-CertificateRevocationList -Force
	New-Alias -Name Get-CTL						-Value Get-CertificateTrustList -Force
	New-Alias -Name Show-CTL					-Value Show-CertificateTrustList -Force
	$aliases += "oid", "oid2", "Get-CRL", "Show-CRL", "Get-CTL", "Show-CTL"
}

# define restricted functions
$RestrictedFunctions =		"Get-RequestRow",
							"__RestartCA",
							"Test-XCEPCompat",
							"Ping-CA",
							"Ping-WMI",
							"Ping-ICertAdmin",
							"Write-ErrorMessage"
$NoDomainExcludeFunctions =	"Add-CAKRACertificate",
							"Add-CATemplate",
							"Add-CertificateEnrollmentPolicyService",
							"Add-CertificateEnrollmentService",
							"Add-CertificateTemplateAcl",
							"Disable-KeyRecoveryAgentFlag",
							"Enable-KeyRecoveryAgentFlag",
							"Get-ADKRACertificate",
							"Get-CAExchangeCertificate",
							"Get-CAKRACertificate",
							"Get-CATemplate",
							"Get-CertificateTemplate",
							"Get-CertificateTemplateAcl",
							"Get-EnrollmentServiceUri",
							"Get-KeyRecoveryAgentFlag",
							"Remove-CAKRACertificate",
							"Remove-CATemplate",
							"Remove-CertificateTemplate",
							"Remove-CertificateTemplateAcl",
							"Restore-KeyRecoveryAgentFlagDefault",
							"Set-CAKRACertificate",
							"Set-CATemplate",
							"Set-CertificateTemplateAcl",
							"Get-CertificationAuthority"
$Win2003ExcludeFunctions =	"Add-CertificateEnrollmentPolicyService",
							"Add-CertificateEnrollmentService",
							"Install-CertificationAuthority",
							"Remove-CertificateEnrollmentPolicyService",
							"Remove-CertificateEnrollmentService",
							"Uninstall-CertificationAuthority"	
$Win2008ExcludeFunctions =	"Add-CertificateEnrollmentPolicyService",
							"Add-CertificateEnrollmentService",
							"Remove-CertificateEnrollmentPolicyService",
							"Remove-CertificateEnrollmentService"
$Win2012ExcludeFunctions =	"Install-CertificationAuthority",
							"Uninstall-CertificationAuthority",
							"Add-CertificateEnrollmentPolicyService",
							"Add-CertificateEnrollmentService",
							"Remove-CertificateEnrollmentPolicyService",
							"Remove-CertificateEnrollmentService"

if ($Win2003) {$RestrictedFunctions += $Win2003ExcludeFunctions}
if ($Win2008) {$RestrictedFunctions += $Win2008ExcludeFunctions}
if ($Win2012) {$RestrictedFunctions += $Win2012ExcludeFunctions}
if ($NoDomain) {$RestrictedFunctions += $NoDomainExcludeFunctions}

# export module members
Export-ModuleMember –Function @(
	Get-ChildItem $PSScriptRoot -Include *.ps1 -Recurse | `
		ForEach-Object {$_.Name -replace ".ps1"} | `
		Where-Object {$RestrictedFunctions -notcontains $_}
)
Export-ModuleMember -Alias $aliases

# stub for types and formats (PS V3+)
if ($PSVersionTable["PSVersion"].Major -gt 2) {
	try {
		Update-TypeData $PSScriptRoot\Types\PSPKI.Types.ps1xml
		Update-FormatData $PSScriptRoot\Types\PSPKI.Format.ps1xml
	} catch { }
}
#endregion
# SIG # Begin signature block
# MIIX1gYJKoZIhvcNAQcCoIIXxzCCF8MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBF94Xc/kTJQhmi
# 3wKEUp0Br6ToV2U/buHHuBlxQ/ZVoqCCEuQwggPuMIIDV6ADAgECAhB+k+v7fMZO
# WepLmnfUBvw7MA0GCSqGSIb3DQEBBQUAMIGLMQswCQYDVQQGEwJaQTEVMBMGA1UE
# CBMMV2VzdGVybiBDYXBlMRQwEgYDVQQHEwtEdXJiYW52aWxsZTEPMA0GA1UEChMG
# VGhhd3RlMR0wGwYDVQQLExRUaGF3dGUgQ2VydGlmaWNhdGlvbjEfMB0GA1UEAxMW
# VGhhd3RlIFRpbWVzdGFtcGluZyBDQTAeFw0xMjEyMjEwMDAwMDBaFw0yMDEyMzAy
# MzU5NTlaMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAsayzSVRLlxwS
# CtgleZEiVypv3LgmxENza8K/LlBa+xTCdo5DASVDtKHiRfTot3vDdMwi17SUAAL3
# Te2/tLdEJGvNX0U70UTOQxJzF4KLabQry5kerHIbJk1xH7Ex3ftRYQJTpqr1SSwF
# eEWlL4nO55nn/oziVz89xpLcSvh7M+R5CvvwdYhBnP/FA1GZqtdsn5Nph2Upg4XC
# YBTEyMk7FNrAgfAfDXTekiKryvf7dHwn5vdKG3+nw54trorqpuaqJxZ9YfeYcRG8
# 4lChS+Vd+uUOpyyfqmUg09iW6Mh8pU5IRP8Z4kQHkgvXaISAXWp4ZEXNYEZ+VMET
# fMV58cnBcQIDAQABo4H6MIH3MB0GA1UdDgQWBBRfmvVuXMzMdJrU3X3vP9vsTIAu
# 3TAyBggrBgEFBQcBAQQmMCQwIgYIKwYBBQUHMAGGFmh0dHA6Ly9vY3NwLnRoYXd0
# ZS5jb20wEgYDVR0TAQH/BAgwBgEB/wIBADA/BgNVHR8EODA2MDSgMqAwhi5odHRw
# Oi8vY3JsLnRoYXd0ZS5jb20vVGhhd3RlVGltZXN0YW1waW5nQ0EuY3JsMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIBBjAoBgNVHREEITAfpB0wGzEZ
# MBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMTANBgkqhkiG9w0BAQUFAAOBgQADCZuP
# ee9/WTCq72i1+uMJHbtPggZdN1+mUp8WjeockglEbvVt61h8MOj5aY0jcwsSb0ep
# rjkR+Cqxm7Aaw47rWZYArc4MTbLQMaYIXCp6/OJ6HVdMqGUY6XlAYiWWbsfHN2qD
# IQiOQerd2Vc/HXdJhyoWBl6mOGoiEqNRGYN+tjCCBKMwggOLoAMCAQICEA7P9DjI
# /r81bgTYapgbGlAwDQYJKoZIhvcNAQEFBQAwXjELMAkGA1UEBhMCVVMxHTAbBgNV
# BAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydTeW1hbnRlYyBUaW1l
# IFN0YW1waW5nIFNlcnZpY2VzIENBIC0gRzIwHhcNMTIxMDE4MDAwMDAwWhcNMjAx
# MjI5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29y
# cG9yYXRpb24xNDAyBgNVBAMTK1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2Vydmlj
# ZXMgU2lnbmVyIC0gRzQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCi
# Yws5RLi7I6dESbsO/6HwYQpTk7CY260sD0rFbv+GPFNVDxXOBD8r/amWltm+YXkL
# W8lMhnbl4ENLIpXuwitDwZ/YaLSOQE/uhTi5EcUj8mRY8BUyb05Xoa6IpALXKh7N
# S+HdY9UXiTJbsF6ZWqidKFAOF+6W22E7RVEdzxJWC5JH/Kuu9mY9R6xwcueS51/N
# ELnEg2SUGb0lgOHo0iKl0LoCeqF3k1tlw+4XdLxBhircCEyMkoyRLZ53RB9o1qh0
# d9sOWzKLVoszvdljyEmdOsXF6jML0vGjG/SLvtmzV4s73gSneiKyJK4ux3DFvk6D
# Jgj7C72pT5kI4RAocqrNAgMBAAGjggFXMIIBUzAMBgNVHRMBAf8EAjAAMBYGA1Ud
# JQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDBzBggrBgEFBQcBAQRn
# MGUwKgYIKwYBBQUHMAGGHmh0dHA6Ly90cy1vY3NwLndzLnN5bWFudGVjLmNvbTA3
# BggrBgEFBQcwAoYraHR0cDovL3RzLWFpYS53cy5zeW1hbnRlYy5jb20vdHNzLWNh
# LWcyLmNlcjA8BgNVHR8ENTAzMDGgL6AthitodHRwOi8vdHMtY3JsLndzLnN5bWFu
# dGVjLmNvbS90c3MtY2EtZzIuY3JsMCgGA1UdEQQhMB+kHTAbMRkwFwYDVQQDExBU
# aW1lU3RhbXAtMjA0OC0yMB0GA1UdDgQWBBRGxmmjDkoUHtVM2lJjFz9eNrwN5jAf
# BgNVHSMEGDAWgBRfmvVuXMzMdJrU3X3vP9vsTIAu3TANBgkqhkiG9w0BAQUFAAOC
# AQEAeDu0kSoATPCPYjA3eKOEJwdvGLLeJdyg1JQDqoZOJZ+aQAMc3c7jecshaAba
# tjK0bb/0LCZjM+RJZG0N5sNnDvcFpDVsfIkWxumy37Lp3SDGcQ/NlXTctlzevTcf
# Q3jmeLXNKAQgo6rxS8SIKZEOgNER/N1cdm5PXg5FRkFuDbDqOJqxOtoJcRD8HHm0
# gHusafT9nLYMFivxf1sJPZtb4hbKE4FtAC44DagpjyzhsvRaqQGvFZwsL0kb2yK7
# w/54lFHDhrGCiF3wPbRRoXkzKy57udwgCRNx62oZW8/opTBXLIlJP7nPf8m/PiJo
# Y1OavWl0rMUdPH+S4MO8HNgEdTCCBRMwggP7oAMCAQICEAGfcm2O2qyxDgPgWB72
# KpowDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTAeFw0xNTEyMTgw
# MDAwMDBaFw0xNjEyMjIxMjAwMDBaMFAxCzAJBgNVBAYTAkxWMQ0wCwYDVQQHEwRS
# aWdhMRgwFgYDVQQKEw9TeXNhZG1pbnMgTFYgSUsxGDAWBgNVBAMTD1N5c2FkbWlu
# cyBMViBJSzCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAOhRW+I+23Aa
# e7xYARsDbO9iPf54kvGula1yiS/JkAsR3yF/ubX3IIiu4KEHdvcKzO04yOBX5rgy
# g80SMx2dsVWy076cLFuH8nVboCuOoQhphfofhkk3B8UPtLbYk14odbv9n/+N2w9J
# NG9K6Ba4YXOLHQPF19MMBO6rXQnqK+LVOT0Nkmkx8QoyfPrN7bhR8lQVfVfFxt4O
# BN0rad3VEYAwqfFhCGfgbO/5Otsslaz3vpotH+0ny13hSq2Ur8ETQ8FLcbtdvh02
# Obh7WdUXPsU1/oOpBDfhkOT5eBVVAg3E1sHZaaQ4wQkVfYbf4Xnf13hXoR9EAXT6
# /VT05+bWbpMCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5
# LfZldQ5YMB0GA1UdDgQWBBT/cEXZqgVC/msreM/XBbwjW3+6gzAOBgNVHQ8BAf8E
# BAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAz
# oDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEu
# Y3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBz
# Oi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4
# MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEF
# BQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFz
# c3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcN
# AQELBQADggEBAFGo/QXI8xd2YZ/gL65sh4dJ4VFy6dLqQV3KiSfy0oocWoC95rxA
# KZ0Wow9NN63RYr/Y7xGKKxxYAMNubIdML0ow06595pta00JvDBoF6DTGKvx6jZ15
# fUlVZ+OLhl3AdOWolHmGcIz6LWIPrTNY7Hv7xYAXq2gKzk7X4IOq3k+G+/RF7RjX
# sN4VZ7001qc53L+35ylO4lmZfdNHl2FFklMxlmdN3OLipNYgBpFfib99R6Ep8HB3
# mnOhnCVnREL/lGdEyl1S1qeTAo92tKMs9I5snAPDGhm9nCkAqHCbXBrj1G/VseD+
# vT3QisKWcBQDo6zU8kBhFYxTxrIwxC4zj3owggUwMIIEGKADAgECAhAECRgbX9W7
# ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBa
# Fw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/l
# qJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fT
# eyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqH
# CN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+
# bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLo
# LFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIB
# yTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAK
# BggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwA
# AgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAK
# BghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0j
# BBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7s
# DVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGS
# dQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6
# r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo
# +MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qz
# sIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHq
# aGxEMrJmoecYpJpkUe8xggRIMIIERAIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEw
# LwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENB
# AhABn3JtjtqssQ4D4Fge9iqaMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcC
# AQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYB
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIITbqHndHLdk
# /R86Xd56SB/yl1SG3g9oLVdpyFqBf8sJMA0GCSqGSIb3DQEBAQUABIIBAGkyLt0Z
# qi1X48xcySh33XKuRyZaS1R1q+YzL2Fs5hZIjsdWpt9Vp4eqj/U02BCms6dKT5vA
# mhFWdBljcim2JhnoeG99HUfh0hdNUmIHkvPOH/EXcNPUJNRJlNlXFN3J6fJBLY2w
# dMU/RdpjmDZ1J61wEMt1J9RdouYY/aaYbpCghOWH/0fFmC3KORqsODMxETyjplQn
# LxvskketDxhuauOB1hUQn+/I0bdcCqFIXDggTD2JOhmPmgoNRcniurJaiyIDmSkk
# GK8gKavX6lkWI90ETDYYOlWfdUMQRAvCMOCn/JNVsRIRBHZYuBcXzsga3DCCLKUm
# ycyUV8TTVIaXQ72hggILMIICBwYJKoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBeMQsw
# CQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNV
# BAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0
# OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwODA3MTgwMDQ5WjAjBgkqhkiG9w0BCQQx
# FgQUvZNF5meF8Bifsi78boeuX5Lh0Z0wDQYJKoZIhvcNAQEBBQAEggEAMTVoA2nX
# 7eRulOfOK8Iumomhe46h03VhHi6Ly6zEZ7nhrIwRVXvzxB5mowcFitlDDWnzwKR0
# UMCvRuko9/pDehI/eNmH62nxdaiBYXpBnQgdBW5IZk1Vr9VdExJLHQWhR/gW1PoA
# fAgZIiZCKXnVVyEcSSTbJgag0YLGyvC1uFnTVRULTAnIuGldnarKoMOdk4dJVMrW
# miPN3PaYNI2RYW3mtjzrRT9lUzYmUR6l7aqqlX5yWWqspPZG7mVE2GVmSGmPFNoV
# LNXFTLufRf3sOv5k/8vDhCxYYajp3FA1ul/B1b9s38Ki5vc+9LwlQ42ihx3c9OHa
# 6AO55pWbLvLfKA==
# SIG # End signature block
