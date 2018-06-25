Set-Alias -Name top -Value Get-TopProcesses

Function Get-TopProcesses {
<#
.SYNOPSIS
    
.DESCRIPTION
    
.EXAMPLE

#>

    Set-Variable -Name SetValX -Value $([Console]::CursorLeft)
    Set-Variable -Name SetValY -Value $([Console]::CursorTop)

    While ($true) {
        Get-Process | Sort-Object -Descending CPU | Select-Object -First 30;
        Start-Sleep -Seconds 2;
        [Console]::SetCursorPosition(${SetValX},${SetValY}+3)
    }
}