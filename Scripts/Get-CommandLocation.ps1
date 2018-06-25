Set-Alias -Name which -Value Get-CommandLocation

Function Get-CommandLocation {

<#
.SYNOPSIS

.DESCRIPTION

.EXAMPLE

#>
    Set-Variable -Name Command

    Param(
        [String]$Command
    )
    
    Get-Command -Name $Command | Select-Object -ExpandProperty Definition
}
