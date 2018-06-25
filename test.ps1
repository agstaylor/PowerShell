# openssl req -modulus -in C:\Users\ataylor\Desktop\request_2017-01-19T16-19-04Z.p10 -noout |  % { $_ -replace “Modulus=”,”” }

param
(
    [Parameter(ValueFromPipeline=$true)] $InputObject
)

    $HexString = @($Input); $HexString = $HexString[0]
    $count     = $HexString.length
    $byteCount = $count/2

    $bytes = New-Object byte[] $byteCount

    for ( $i = 0; $i -le $count-1; $i+=2 )
    { 
        $bytes[$i/2] = [byte]::Parse($HexString.Substring($i,2), [System.Globalization.NumberStyles]::HexNumber)
    }

    $OutputFile = "binary.bin"
    set-content -encoding byte $OutputFile -value $bytes

 