## Environment variables go here
Set-Variable -Name Profile_Directory -Value $(Split-Path -Path ${myInvocation}.MyCommand.Path -Parent) -Scope Private
Set-Item -Path Env:\HOMEDRIVE -Value $("$((Get-Item ${Profile_directory}).PSDrive.Name):") -Force
Set-Item -Path Env:\HOMEPATH -Value $(Split-Path ((Get-Item ${Profile_Directory}).Parent.Parent.FullName) -NoQualifier) -Force
Set-Item -Path Env:\USERPROFILE -Value $((Get-Item ${Profile_Directory}).Parent.Parent.FullName) -Force
Set-Item -Path Env:\APPDATA -Value $(Join-Path -Path ${env:USERPROFILE} -ChildPath Appdata\Roaming -Resolve) -Force
Set-Item -Path Env:\LOCALAPPDATA -Value $(Join-Path -Path ${env:USERPROFILE} -ChildPath Appdata\Local -Resolve) -Force
Set-Item -Path Env:\TEMP -Value $(Join-Path -Path ${env:LOCALAPPDATA} -ChildPath Temp) -Force
Set-Item -Path Env:\TMP -Value $(${env:TEMP}) -Force
Set-Item -Path Env:\PSModulePath -Value $(${env:PSModulePath} + ";${Profile_Directory}\Modules") -Force
Set-Item -Path Env:\PATH -Value $(${env:PATH} + ";${Profile_Directory}\Stash\Bin64" + ";${Profile_Directory}\Stash\Bin") -Force

## Script variables go here
    # Check for Administrator elevation
Set-Variable -Name UserInfo -Value $([System.Security.Principal.WindowsIdentity]::GetCurrent())  -Scope Private
Set-Variable -Name CheckGroup -Value $(New-Object -TypeName System.Security.Principal.WindowsPrincipal -ArgumentList (${UserInfo}))  -Scope Private
Set-Variable -Name AdminPrivCheck -Value $([System.Security.Principal.WindowsBuiltInRole]::Administrator) -Scope Private
Set-Variable -Name AdminPrivSet -Value $(${CheckGroup}.IsInRole(${AdminPrivCheck}))
Set-Variable -Name MyHostName -Value $([System.Net.Dns]::GetHostName())

## Functions and other gibblets here
    # Print the user and host:
Function Prompt {

        # Check for Administrator elevation
    If (${AdminPrivSet}) {
        Write-Host -Object ("${env:UserName}") -NoNewline -ForegroundColor Red

    } else {
        Write-Host -Object ("${env:UserName}") -NoNewline -ForegroundColor Green
    }

    Write-Host -Object ("@") -NoNewline -ForegroundColor Green
    #Write-Host -Object ("${env:ComputerName}") -NoNewline -ForegroundColor Green
    Write-Host -Object ("$MyHostName") -NoNewline -ForegroundColor Green
    
        # Print the working directory:
    Write-Host -Object (":") -NoNewline -ForegroundColor Cyan
        # Replace string matching homedir with tilde
    Write-Host -Object ($((Get-Location).Path).Replace(${HOME},"~")) -NoNewLine -ForegroundColor Cyan

    If (${AdminPrivSet}) {
        Write-Host -Object (">") -NoNewline -ForegroundColor Red
        return " "
    }
    else {
        Write-Host -Object (">") -NoNewline -ForegroundColor White
        return " "
    }
}

    # Import user scripts
Set-Variable -Name Script_Dir -Value $(Join-Path -Path ${Profile_Directory} -ChildPath "Scripts") -Scope Private

Get-ChildItem -Path ${Script_Dir} | Where-Object {
    $_.Name -like "*.ps1"
} | ForEach-Object {
    . $_.FullName
}

## cd $my_dirs.nfast
$my_dirs = @{
       nfast = 'C:\Program Files (x86)\nCipher\nfast\bin'
    projects = 'D:\Projects'
}

# use GIT curl, not PS alias
remove-item alias:curl

# python virtual envs
Import-module VirtualEnvWrapper

## Say hello
Get-MOTD
