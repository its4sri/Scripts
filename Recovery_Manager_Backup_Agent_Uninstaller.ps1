[CmdletBinding()]
 
param (
 
[parameter(Mandatory=$true,Position=1)]
[string[]]$computer
 
 
)

ForEach ($c in $computer){
 
    #set script execution on remote machine
    $bypass = invoke-command -ComputerName $c -ScriptBlock {set-executionpolicy bypass -Force}

    #use TLS1.2 on remote machine
    $tls12 = invoke-command -ComputerName $c -ScriptBlock {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12}

    #update nuget module on remote machine
    #$nugetinstall = invoke-command -ComputerName $c -ScriptBlock {Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force}
    
        #Reset Timeouts
        $connectiontimeout = 0
                
        #starts up a remote powershell session to the computer
        do{
            $session = New-PSSession -ComputerName $c
            "reconnecting remotely to $c"
            sleep -seconds 10
            $connectiontimeout++
        } until ($session.state -match "Opened" -or $connectiontimeout -ge 10)
 
            Write "Uninstalling..."
            invoke-command -session $session -scriptblock {start-process "MsiExec.exe /quiet /x {40582CF1-C31D-4319-920C-2D8CCA02EB32}" -Wait}
 
            sleep -Seconds 30
 
            #restarts the remote computer and waits till it starts up again
 
            Write "restarting remote computer"
 
            # maybe create a log file using date
            # $date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
 
            Restart-Computer -Wait -ComputerName $c -Force

            # We then then create a new seesion connection and do the copy on of new agent and install once the server has rebooted.....  
 
}
