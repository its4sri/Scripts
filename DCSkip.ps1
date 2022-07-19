$ServerFileData = Get-Content d:\scripts\decommlist.txt
foreach ($Server in $ServerFileData)
{
    checkForDC($Server)
   #echo "inside decomtest script"
}
	
function checkForDC( [string] $arg )
  {
 
 $server,$domain = $arg.split('.')
 $String = Get-ADComputer -SearchBase "OU=File and Print,OU=AVZ Servers,DC=corpdev,dc=dev" -filter * | Select  Name
 #$CORPDCList = Get-WmiObject -Class Win32_Desktop -ComputerName . | Select Name

#$String = $CORPDEVDCList + $CORPDCList

  If ($String -match $server)
   { 
        Write-Output 'The given input server '$arg' is present in the Domain Controllers hence skipping the rebooting part'
		Write-Output 'removing from Reboot machine list file'
		removeFromReboot($arg)
    }
    Else 
    {
     Write-Output 'The given input server '$arg' is not Domain Controlller. Hence going for rebooting' 
     #$psfile =  "C:\Users\uppulsk\Downloads\abc.ps1"
     #.$psfile 
     
     }

  }
  
 function removeFromReboot( [string] $arg )
 {
 $ServerFileData = Get-Content d:\scripts\decommlist.txt
 #echo $arg
 $del = "$arg"
 $ServerFileData = $ServerFileData | Where {$_ -ne $del}
 $ServerFileData | Out-File d:\scripts\decommlist.txt -Force