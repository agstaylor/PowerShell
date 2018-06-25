Set-Alias file Get-MimeType

function Get-MimeType {
<#
.SYNOPSIS
    
.DESCRIPTION
    
.EXAMPLE
    PS > Get-MimeType 'C:\Users\Foo\bar.mp3'
    C:\Users\Foo\bar.mp3:audio/mpeg

#>

    param(
        [Parameter(Mandatory=$true)]
        [String]$FileName
    )
  
    begin {
        $ext = ([IO.FileInfo](cvpa $FileName)).Extension.ToLower()
        $res = 'application/unknown'
    }
    process {
        try {
            $rk = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($ext)
        }
    finally {
        if ($rk -ne $null) {
            if (($cur = $rk.GetValue('Content Type')) -ne $null) {
            $res = $cur
            }
            $rk.Close()
        }
        }
    }
    end {
        Write-Host $FileName`: -f Yellow -no
        $res
    }
}
