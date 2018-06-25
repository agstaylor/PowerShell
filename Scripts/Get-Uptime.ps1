Set-Alias -Name uptime -Value Get-Uptime

Function Get-Uptime([String]${Computer} = '.') {
<#
.SYNOPSIS
    
.DESCRIPTION
    
.EXAMPLE

#>

    Set-Variable -Name OSObject -Value $(Get-WmiObject -Class Win32_OperatingSystem)
    Set-Variable -Name Uptime -Value $((Get-Date) - (${OSObject}.ConvertToDateTime(${OSObject}.lastbootuptime)))
    Set-Variable -Name Output -Value $("Uptime: " + ${Uptime}.Days + " days, " + ${Uptime}.Hours + " hours, " + ${Uptime}.Minutes + " minutes")
    Write-Output -InputObject ${Output}

}
