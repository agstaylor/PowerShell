function Get-WebFile {

<#
.SYNOPSIS
    Downloads a file or page from the web

.DESCRIPTION
    
.EXAMPLE

#>

    param( 
        $url = (Read-Host "The URL to download"),
        $fileName = $null,
        [switch]$Passthru,
        [switch]$quiet
    )
   
    $req = [System.Net.HttpWebRequest]::Create($url);
    $res = $req.GetResponse();
 
    if($fileName -and !(Split-Path $fileName)) {
        $fileName = Join-Path (Get-Location -PSProvider "FileSystem") $fileName
    } 
    elseif((!$Passthru -and ($fileName -eq $null)) -or (($fileName -ne $null) -and (Test-Path -PathType "Container" $fileName))) {
        [string]$fileName = ([regex]'(?i)filename=(.*)$').Match( $res.Headers["Content-Disposition"] ).Groups[1].Value
        $fileName = $fileName.trim("\/""'")
        if(!$fileName) {
            $fileName = $res.ResponseUri.Segments[-1]
            $fileName = $fileName.trim("\/")
            if(!$fileName) { 
                $fileName = Read-Host "Please provide a file name"
            }
            $fileName = $fileName.trim("\/")
            if(!([IO.FileInfo]$fileName).Extension) {
                $fileName = $fileName + "." + $res.ContentType.Split(";")[0].Split("/")[1]
            }
        }
        $fileName = Join-Path (Get-Location -PSProvider "FileSystem") $fileName
    }
	
    if($Passthru) {
        $encoding = [System.Text.Encoding]::GetEncoding( $res.CharacterSet )
        [string]$output = ""
    }
 
    if($res.StatusCode -eq 200) {
        [int]$goal = $res.ContentLength
        $reader = $res.GetResponseStream()
        if($fileName) {
            $writer = new-object System.IO.FileStream $fileName, "Create"
        }
        [byte[]]$buffer = new-object byte[] 4096
        [int]$total = [int]$count = 0
        do {
            $count = $reader.Read($buffer, 0, $buffer.Length);
            if($fileName) {
                $writer.Write($buffer, 0, $count);
            } 

            if($Passthru){
                $output += $encoding.GetString($buffer,0,$count)
            } 
            elseif(!$quiet) {
                $total += $count
                if($goal -gt 0) {
                    Write-Progress "Downloading $url" "Saving $total of $goal" -id 0 -percentComplete (($total/$goal)*100)
                }
                else {
                    Write-Progress "Downloading $url" "Saving $total bytes..." -id 0
                }
            }
        }
		while ($count -gt 0)
      
        $reader.Close()
        if($fileName) {
            $writer.Flush()
            $writer.Close()
        }

        if($Passthru){
            $output
        }
    }

    $res.Close(); 
    if($fileName) {
        ls $fileName
    }
}
