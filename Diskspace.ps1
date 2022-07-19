#######################################################################  
# This PS script provides C drive free spaces.
####################################################################### 
$Server = Get-Content "\\ushoudsutl01\scripts\Diskspace\servers.txt"
Get-WMIObject Win32_LogicalDisk -ComputerName $Server| Where-Object {$_.deviceid -eq "c:"} | Select-Object @{name="ServerName";expression={$_.PSComputerName}}, @{name="Drive";expression={$_.deviceid}}, @{name="FreeSpace(GB)";expression={$_.freespace / 1GB}}, @{name="TotalSize(GB)";expression={$_.Size / 1GB}} | export-csv \\ushoudsutl01\scripts\Diskspace\servers.csv 