$ComputerNames = Get-Content D:\scripts\CheckServer\Servers.txt
Function CheckServer {

    begin {
        $SelectHash = @{
         'Property' = @('Server Name','ServerInAD','ServerInDNS','ResponseToPing','SharesAccessible','RDPAccessible','Uptime')
        }
    }

    process {
        foreach ($CurrentComputer in $ComputerNames) {
        $DefUptime = "0 Days 0 Hours 0 Min 0 Sec"
# Create new Hash
            $HashProps = @{
                'Server Name' = $CurrentComputer
                'ServerInAD' = $false
                'ServerInDNS' = $false
                'ResponseToPing' = $false
                'SharesAccessible' = $false
                'RDPAccessible' = $false
                'Uptime' = $DefUptime
            }
        
            # Perform Checks
            switch ($true)
            {
                {([adsisearcher]"samaccountname=$CurrentComputer`$").findone()} {$HashProps.ServerInAD = $true}
                {$(try {[system.net.dns]::gethostentry($CurrentComputer)} catch {})} {$HashProps.ServerInDNS = $true}
                {Test-Connection -ComputerName $CurrentComputer -Quiet -Count 1} {$HashProps.ResponseToPing = $true}
                {get-WmiObject -class Win32_Share -computer $CurrentComputer} {$HashProps.SharesAccessible = $true}
                {$(try {$socket = New-Object Net.Sockets.TcpClient($CurrentComputer, 3389);if ($socket.Connected) {$true};$socket.Close()} catch {})} {$HashProps.RDPAccessible = $true}
                Default {}
            }
           $Computerobj = "" | select ComputerName, Uptime, LastReboot 
           $wmi = Get-WmiObject -ComputerName $CurrentComputer -Query "SELECT LastBootUpTime FROM Win32_OperatingSystem" -ErrorAction SilentlyContinue
           $now = Get-Date
                if (!($wmi -eq $null)) {
           $boottime = $wmi.ConvertToDateTime($wmi.LastBootUpTime)
           $uptime = $now - $boottime
           $d =$uptime.days
           $h =$uptime.hours
           $m =$uptime.Minutes
           $s = $uptime.Seconds
           $Computerobj.ComputerName = $CurrentComputer
           $Computerobj.Uptime = "$d Days $h Hours $m Min $s Sec"
                $Computerobj.LastReboot = $boottime
           $HashProps.Uptime = $Computerobj.Uptime
                }

            # Output object
            New-Object -TypeName 'PSCustomObject' -Property $HashProps | Select-Object @SelectHash
        }
    }

    end {
    }
}
#CheckServer | out-host 
CheckServer | out-file D:\scripts\CheckServer\CheckServerResults.txt
