
Function Test-Port ($DestinationHosts,$Ports,[switch]$noPing,$pingTimeout="2000",[switch]$ShowDestIP,[Switch]$Continuous,$waitTimeMilliseconds=300) { 
  
    
    #ScriptBlock to check the port and return the result
    $Script_CheckPort = {
        $Source = $Args[0];
        $Destination = $args[1];
        $Port = $args[2];
        $PingTimeout = $Args[3]; 
        $ShowIP = $args[4];

        #Generate the result Object
        $result = New-Object psobject -property @{Source=$Source;Dest=$destination;Port=$Port;Result="";Error=$()}

        #If IP is requested add to the result
        if ($ShowIP){
            try {
                $result | add-member -MemberType NoteProperty -Name "Dest IP" -value (([System.Net.Dns]::GetHostAddresses($Destination) | select -expand IPAddressToString) -join ";") -Force
            } catch {
                $result.error += $_;
                $result | add-member -MemberType NoteProperty -Name "Dest IP" -value $Destination -Force
            }
        }

        #if it is a Ping
        if ($Port -eq "PING"){
            try {
                $ping = new-object System.Net.NetworkInformation.Ping
                $result.Result = $ping.send($destination,$PingTimeout).Status
            } catch {
                $result.error += $_;
                $result.Result = "Unresolvable"
            } finally {
                if($ping){$ping.dispose();$ping = $null}
            }
        
        } else {
        #it is a socket
            try{
                $socket = new-object System.Net.Sockets.TcpClient($destination, $port)
                if ($socket -eq $null) {
                    
                } elseif ($socket.connected -eq $true) {
                    $result.result = "Success"
                }
            } catch {
                $result.error += $_
                $thisError = $_
                switch -regex ($_.ToString()) {
                    'actively refused' { $result.result ='Refused'; break;}
                    'No such host is known' {$result.result = 'HostName Error';break;}
                    default {$result.result = 'FAIL'}
                }
            } finally {
                if ($socket.Connected){
                    $socket.close()
                }
                if ($socket){
                    $socket.Dispose()
                }
                $socket = $null
            }
        }
        Write-Output $result

    }
    #Get the list of ports for the query
    if (!$noPing) {$portsToQuery = @('PING');} else {$portsToQuery=@()}

     switch -regex ($ports) {
        '(?i)Domain|DC|AD|Active Directory' {$portsToQuery += @(445,389,88,3268)}
        '(?i)Web' {$portsToQuery += @(80,443)}
        '(?i)http(^s)' {$portsToQuery += 80 }
        '(?i)https' {$portsToQuery += 443 }
        '(?i)smb' {$portsToQuery += 445 }
        '(?i)SCOM|Ops Man|Operations Manager' {$portsToQuery += 5723}
        '(?i)rdp|remote desktop|mstsc' {$portsToQuery += 3389}
        '(?i)remoting|[^t]PS|PowerShell' {$portsToQuery += @(5985,5986)}
        '(?i)eset' {$portsToQuery += @(2222,2221)}
        default {if ($_ -is [int32]){$portsToQuery+=$_}}     
     }
     $PortsToQuery = $POrtsToQuery | Sort-Object -descending | Select-object -unique
     $hostname = hostname
     #run the lookup

     $maxthreads = 5;
     $iss = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
     $pool = [Runspacefactory]::CreateRunspacePool(1, $maxthreads, $iss, $host)
     $pool.open()
     $threads = @()
     $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock($Script_CheckPort.toString())
     $FirstRun = $True
     try {           
         While ($continuous -or $FirstRun) {
             foreach ($Destination in $DestinationHosts) {
                foreach ($port in $POrtsToQuery) {
                    $powershell = [powershell]::Create().addscript($scriptblock).addargument($hostname).addargument($Destination).AddArgument($port).AddArgument($pingTimeout).AddArgument($ShowDestIP)
                    $powershell.runspacepool=$pool
                    $threads+= @{
                        instance = $powershell
                        handle = $powershell.begininvoke()
                    }            
           
                }
             }
     
             $notdone = $true
             while ($notdone) {
                $notdone = $false
                for ($i=0; $i -lt $threads.count; $i++) {
                    $thread = $threads[$i]
                    if ($thread) {
                        if ($thread.handle.iscompleted) {
                            $thread.instance.endinvoke($thread.handle)
                            $thread.instance.dispose()
                            $threads[$i] = $null
                        }
                        else {
                            $notdone = $true
                        }
                    }
                }
                Start-Sleep -milliseconds 300
            }
            if ($continuous) {
                Start-Sleep -Milliseconds $waitTimeMilliseconds
            } else {
                $firstrun = $False
            }
        }
    } catch {
        throw $_
    } finally {
        $pool.close()  
        $iss = $null
        $threads = $null
        $scriptblock = $null
    }

}

Write-Output "hello"
Test
testnew-alias tp Test-Port 
#Write-HOst "Test-Port Alias:tp"
Function Test-Port ($DestinationHosts,$Ports,[switch]$noPing,$pingTimeout="2000",[switch]$ShowDestIP,[Switch]$Continuous,$waitTimeMilliseconds=300) { 
  
    
    #ScriptBlock to check the port and return the result
    $Script_CheckPort = {
        $Source = $Args[0];
        $Destination = $args[1];
        $Port = $args[2];
        $PingTimeout = $Args[3]; 
        $ShowIP = $args[4];

        #Generate the result Object
        $result = New-Object psobject -property @{Source=$Source;Dest=$destination;Port=$Port;Result="";Error=$()}

        #If IP is requested add to the result
        if ($ShowIP){
            try {
                $result | add-member -MemberType NoteProperty -Name "Dest IP" -value (([System.Net.Dns]::GetHostAddresses($Destination) | select -expand IPAddressToString) -join ";") -Force
            } catch {
                $result.error += $_;
                $result | add-member -MemberType NoteProperty -Name "Dest IP" -value $Destination -Force
            }
        }

        #if it is a Ping
        if ($Port -eq "PING"){
            try {
                $ping = new-object System.Net.NetworkInformation.Ping
                $result.Result = $ping.send($destination,$PingTimeout).Status
            } catch {
                $result.error += $_;
                $result.Result = "Unresolvable"
            } finally {
                if($ping){$ping.dispose();$ping = $null}
            }
        
        } else {
        #it is a socket
            try{
                $socket = new-object System.Net.Sockets.TcpClient($destination, $port)
                if ($socket -eq $null) {
                    
                } elseif ($socket.connected -eq $true) {
                    $result.result = "Success"
                }
            } catch {
                $result.error += $_
                $thisError = $_
                switch -regex ($_.ToString()) {
                    'actively refused' { $result.result ='Refused'; break;}
                    'No such host is known' {$result.result = 'HostName Error';break;}
                    default {$result.result = 'FAIL'}
                }
            } finally {
                if ($socket.Connected){
                    $socket.close()
                }
                if ($socket){
                    $socket.Dispose()
                }
                $socket = $null
            }
        }
        Write-Output $result

    }
    #Get the list of ports for the query
    if (!$noPing) {$portsToQuery = @('PING');} else {$portsToQuery=@()}

     switch -regex ($ports) {
        '(?i)Domain|DC|AD|Active Directory' {$portsToQuery += @(445,389,88,3268)}
        '(?i)Web' {$portsToQuery += @(80,443)}
        '(?i)http(^s)' {$portsToQuery += 80 }
        '(?i)https' {$portsToQuery += 443 }
        '(?i)smb' {$portsToQuery += 445 }
        '(?i)SCOM|Ops Man|Operations Manager' {$portsToQuery += 5723}
        '(?i)rdp|remote desktop|mstsc' {$portsToQuery += 3389}
        '(?i)remoting|[^t]PS|PowerShell' {$portsToQuery += @(5985,5986)}
        '(?i)eset' {$portsToQuery += @(2222,2221)}
        default {if ($_ -is [int32]){$portsToQuery+=$_}}     
     }
     $PortsToQuery = $POrtsToQuery | Sort-Object -descending | Select-object -unique
     $hostname = hostname
     #run the lookup

     $maxthreads = 5;
     $iss = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
     $pool = [Runspacefactory]::CreateRunspacePool(1, $maxthreads, $iss, $host)
     $pool.open()
     $threads = @()
     $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock($Script_CheckPort.toString())
     $FirstRun = $True
     try {           
         While ($continuous -or $FirstRun) {
             foreach ($Destination in $DestinationHosts) {
                foreach ($port in $POrtsToQuery) {
                    $powershell = [powershell]::Create().addscript($scriptblock).addargument($hostname).addargument($Destination).AddArgument($port).AddArgument($pingTimeout).AddArgument($ShowDestIP)
                    $powershell.runspacepool=$pool
                    $threads+= @{
                        instance = $powershell
                        handle = $powershell.begininvoke()
                    }            
           
                }
             }
     
             $notdone = $true
             while ($notdone) {
                $notdone = $false
                for ($i=0; $i -lt $threads.count; $i++) {
                    $thread = $threads[$i]
                    if ($thread) {
                        if ($thread.handle.iscompleted) {
                            $thread.instance.endinvoke($thread.handle)
                            $thread.instance.dispose()
                            $threads[$i] = $null
                        }
                        else {
                            $notdone = $true
                        }
                    }
                }
                Start-Sleep -milliseconds 300
            }
            if ($continuous) {
                Start-Sleep -Milliseconds $waitTimeMilliseconds
            } else {
                $firstrun = $False
            }
        }
    } catch {
        throw $_
    } finally {
        $pool.close()  
        $iss = $null
        $threads = $null
        $scriptblock = $null
    }

}

Write-Output "hello"
Test
testnew-alias tp Test-Port 
#Write-HOst "Test-Port Alias:tp"