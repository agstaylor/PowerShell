function Submit-CertificateRequest {
<#
.ExternalHelp PSPKI.Help.xml
#>
[OutputType('PKI.Enrollment.CertRequesStatus')]
[CmdletBinding(DefaultParameterSetName = '__dcom')]
	param(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[IO.FileInfo[]]$Path,
		[Parameter(Mandatory = $true, ParameterSetName = '__dcom')]
		[Alias('CA')]
		[PKI.CertificateServices.CertificateAuthority]$CertificationAuthority,
		[Parameter(Mandatory = $true, ParameterSetName = '__xcep')]
		[Alias('CEP')]
		[PKI.Enrollment.Policy.PolicyServerClient]$EnrollmentPolicyServer,
		[System.Management.Automation.PSCredential]$Credential,
		[String[]]$Attribute
	)
	begin {
		$ErrorActionPreference = "Stop"
		$CertRequest = New-Object -ComObject CertificateAuthority.Request
		switch ($PsCmdlet.ParameterSetName) {
			"__xcep" {
				if (!(Test-XCEPCompat)) {throw New-Object NotSupportedException}
				if (![string]::IsNullOrEmpty($Credential.UserName)) {
					switch ($EnrollmentPolicyServer.Authentication) {
						"UserNameAndPassword" {
							$CertRequest.SetCredential(
								0,
								[int]$EnrollmentPolicyServer.Authentication,
								$Credential.UserName,
								[Runtime.InteropServices.Marshal]::PtrToStringAuto(
									[Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
								)
							)
						}
						"ClientCertificate" {							
							$CertRequest.SetCredential(
								0,
								[int]$EnrollmentPolicyServer.Authentication,
								$Credential.UserName,
								$null
							)
						}
					}
				}
			}
			"__dcom" {
				if (!$CertificationAuthority.IsAccessible) {
					$e = New-Object PKI.Exceptions.ServerUnavailableException
					$e.Server = $CertificationAuthority.DisplayName
					throw $e
				}
			}
		}
		if ($Attribute -eq $null) {
			$strAttribute = [string]::Empty
		} else {
			$SB = New-Object Text.StringBuilder
			foreach ($attrib in $Attribute) {
				[Void]$SB.Append($attrib + "`n")
			}
			$strAttribute = $SB.ToString()
			$strAttribute = $strAttribute.Substring(0,$strAttribute.Length - 1)
		}
	}
	process {
		$Path | ForEach-Object {
			try {
				$Request = [IO.File]::ReadAllText($_)
				$Status = $CertRequest.Submit(0xff,$Request,$strAttribute,$CertificationAuthority.ConfigString)
				$Output = New-Object PKI.Enrollment.CertRequestStatus -Property @{
					CertificationAuthority = $CertificationAuthority;
					Status = $Status;
					RequestID = $CertRequest.GetRequestId()
				}
				if ($Status -eq 3) {
					$base64 = $CertRequest.GetCertificate(1)
					$Output.Certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2 (,[Convert]::FromBase64String($base64))
				} else {
					$Output.ErrorInformation = $CertRequest.GetDispositionMessage()
				}
				$Output
			} catch {throw $_}
		}
	}
}
# SIG # Begin signature block
# MIIX1gYJKoZIhvcNAQcCoIIXxzCCF8MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDBMGBBoc1cDhU
# f0v2iMBYoNCHM51nUwCkof8WzyhfeqCCEuQwggPuMIIDV6ADAgECAhB+k+v7fMZO
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINde+b4Qsuho
# wYVZmLPc3e3vZgdtnMS/+ygLDhJiPAx3MA0GCSqGSIb3DQEBAQUABIIBALHxcoSz
# rwTyFEUGBvROvI2djP3j1ci/i52CdxxIap2XbTcX6oe95/K7n1Jk0OrMM2v2LbB5
# pQP2Jq2NUZptTgFr0xuhCKIz35Gd93M5IhFSHsi35GojfxiwARyYbP8UAqwd8wHX
# mzWJjuEourm14a6fn1q89DbUMi+dLvjeEVBB66UKPQmnhCABsKZ+NyF6ppF8jwkS
# Mtrdy5O+T28HkC6LZzXYRGiEFDBdnNHIX9+jTH1sZE8X/QNOsuyMW47AxepBpUH5
# VWhF4v0nCZIywCVVi1VJIbzy0oSpY5qdcPrziqwIgbV/yGqr9h19nG5Rv3gmz/Dx
# qspjfYeCzpByO1ahggILMIICBwYJKoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBeMQsw
# CQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNV
# BAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0
# OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwODA3MTgwMTUxWjAjBgkqhkiG9w0BCQQx
# FgQUt0qzxh+KkFomKJ+2SY+eK5C77ugwDQYJKoZIhvcNAQEBBQAEggEALP492xFQ
# KJuv8BuBzHDHWdGuAmNJa+m4S22tqgwGlJqkCgz1EyDLT4Sb2nTHvdXQi8ofXGpt
# Yrk9l9lpVwP+veAhq508xFenQXjiqVE+IfYph/G9cdZp4JTAYPeXW1ljpFuDco/E
# vBwQaBgO1Tq9BmIzN8H0UmOauAk4oWTL2PeIvnIuF8FJCug7oXBvHBosDMO/peGp
# vf1MwXBymyVge3SEJki4ubtzZpEsM6cbW2UcJx55UOxSlPPBs8fIXCDjHHoVu2E+
# I6TWoHN4FHIk/ZYMj1sL7hf4LToOetnNlITW7rrrscXgYLIqw83KkY/69UKD7hmq
# bm43rK+VVwN6Kw==
# SIG # End signature block
