
<#

.PARAMETER vc
Name of the virtual center containing the vm

.PARAMETER vm
Name of the virtual machine

.PARAMETER snid
Referenece number from ServiceNow for snapshot

.NOTES
       Change Log    
       2014-12-09 version 1.0
              Intial creation
       
#>

# param (     
#      [Parameter(Mandatory=$true)]
#      [string]$vc,
#      [Parameter(Mandatory=$true)]
#      [string]$vm,
#      [string]$snid 
#      )



Function Cleanup-Variables {
<#
.SYNOPSIS
Function collection
.DESCRIPTION

.PARAMETER Command
Begin - This is executed at the beginning of a script to collect the established variables

End - This is executed when closing a script to remove the variables not seen when the script started

.EXAMPLE
---Example #1---
Cleanup-Variables "Begin"

--Example #2---
Cleanup-Variables "End"

#>
$WhatToDo = $args[0]

Switch ($WhatToDo) {
       "Begin" {
              new-variable -force -name startupVariables -value ( Get-Variable |   % { $_.Name } )
              $LoadedSnapins = Get-PSSnapin
              }
       "End" {       
              Get-PSSnapin | Where-Object { $LoadedSnapins -notcontains $_.Name}| % {Remove-PSSnapin -Name "$($_.Name)" -ErrorAction SilentlyContinue  }
              Get-Variable | Where-Object { $startupVariables -notcontains $_.Name } | % { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue}
              }
       }      
}      

Function Load-Snapin {
<#
.SYNOPSIS
Checks if a snapin is loaded, if not checks for resistration and then loads it

.DESCRIPTION

.PARAMETER Snapin
The name of the snapin to be loaded

.EXAMPLE
Loads the powercli addin from VMWARE

Load-Snapin "VMware.VimAutomation.Core"
#>

$snapin=$args[0]
if (get-pssnapin $snapin -ea "silentlycontinue") {
write-host "PSsnapin $snapin is loaded" -foregroundcolor Blue
}
elseif (get-pssnapin $snapin -registered -ea "silentlycontinue") {
write-host "PSsnapin $snapin is registered but not loaded" -ForegroundColor Yellow -BackgroundColor Black
Add-PSSnapin $snapin
Write-Host "PSsnapin $snapin is loaded" -ForegroundColor Blue
}
else {
write-host "PSSnapin $snapin not found" -foregroundcolor Red
}

}

Function Get-Email ([string]$lanid) {
###Function Get-Email
###This function returns the e-mail address of the logged on users
###If the account is a the corp-svc-ctx-xa account or an autosys service account it returns a specific address
###Otherwise to takes the employee ID of the logged in user and looks for an account in the client OU with that same
###employee ID and returns that as the address
#$username = [Environment]::UserName
$username=$lanid
Switch ($username)
{
"corp-svc-ctx-xa" {return "CORP-SVC-CTX-XA@CORP.AMVESCAP.NET"}
"ushou-asysws*" {return "AutosysSTG@aiminvestments.com" }
"ushou-asyswd*" {return "AutosysDEV@aiminvestments.com" }
"ushou-asyswp*" {return "AutosysACE@aiminvestments.com" }

default {
    $defaultNamingContext=([ADSI]("LDAP://rootDSE")).defaultNamingContext
    $query = "(sAMAccountName=$username)"
    $attrs = @("cn")
    $searcher = New-Object DirectoryServices.DirectorySearcher([ADSI]("LDAP://$defaultNamingContext"), $query, $attrs)
    $objUser = $searcher.FindOne()
    if ($objUser) {
        $emplID = $objUser.GetDirectoryEntry().extensionAttribute1
        $query = "(&(objectCategory=person)(objectClass=user)(extensionAttribute1=$emplID)(mail=*))"
        $searcher = New-Object DirectoryServices.DirectorySearcher([ADSI]("LDAP://OU=AVZ Clients,$defaultNamingContext"), $query, $attrs)
        $objUser = $searcher.FindOne()
        if ($objUser) {
            return $objUser.GetDirectoryEntry().mail
        } else {
            "Regular account not found"
        }
    } else {
        "Admin account not found"
    }
} #closing default switch option
} # closing switch command
} #closing function




############################## 
 
 ###End of Function Defintion(s)###

#$ErrorActionPreference = "SilentlyContinue"

$file = "\\ushoudsutl01\Scripts\SnapshotCreation\Input_Servers.xlsx"
$objExcel=New-Object -ComObject Excel.Application
$objExcel.Visible=$false
$WorkBook=$objExcel.Workbooks.Open($file)
$worksheet = $workbook.sheets.item("SHEET1")
$intRowMax =  ($worksheet.UsedRange.Rows).count
$intColumnMax =  ($worksheet.UsedRange.Columns).count
$Columnnumber = 1
for($intRow = 2 ; $intRow -le $intRowMax ; $intRow++)
{
$ServerName = $worksheet.cells.item($intRow,1).value2 
$SNNumber = $worksheet.cells.item($intRow,2).value2 
$Lanid = $worksheet.cells.item($intRow,3).value2 
$Requestor = $worksheet.cells.item($intRow,4).value2  
$vm,$SNNumber,$lanid,$Requestor
   
If($vm -like '*US*')
{
    $VC='USHOUVC02V'
}
ELSEIf($vm -like '*GB*')
{
    $VC='GBLONVC02V'
}
ELSE
{
    $VC='INHYDVC02V'

}



#write-host $vm,$desc,$snnumber,$lanid,$requestor,$vc
#}

#$VM='USHOUBUILD01VT'

#$Requestor='sahus'

$email=get-email $Requestor
If(Test-Connection -ComputerName $vm)
{
Cleanup-Variables "Begin"
$oldverbose = $VerbosePreference
#$VerbosePreference = "continue"
$VerbosePreference = "silentlycontinue"
#Getting e-mail address
$datestring = (Get-Date).ToString('dd-MM-y')
$snid=$snnumber
Write-Host "ServiceNow record :"$snid       
If ($snid.length -eq 0) {$snid = "1"}  
Write-Verbose $email
Write-Verbose $datestring
Write-Verbose $env:USERNAME
Write-Verbose $snid
Load-Snapin "VMware.VimAutomation.Core"
If ($Defaultviservers -ne $null) {$Defaultviservers | Disconnect-VIServer -Confirm:$False}
Connect-VIServer -server $vc
$vmobject = Get-VM -Name $vm | Where {$_.powerstate -eq "PoweredOn"}

$snapshotlist = Get-Snapshot -VM $vmobject
$createsnapshot = $true

If ($snapshotlist.Count -gt 2) { $createsnapshot = $false 
  $BODY = "The following snapshots already exist:"
       Foreach ($snap in $snapshotlist) {
                     $BODY+= $snap.name + '         ,         '
                     }     
                     $subject='Snapshot Creation skipped for :'+ $vm                                   
      #Send-MailMessage -To $email -Cc "Operations.Server@Invesco.com"  -Subject "Snapshort Creation failed $vmthe new one" -From "MBFN6270@invesco.com" -SmtpServer "emailnasmtp.app.invesco.net" -Body $BODY            
                     }




If ($createsnapshot -eq $true) {
       $snapshotname = "AUTOSNAPSHOT : " + $snid + " : " + $datestring + " : " + $env:USERNAME
       While ($snapshotlist.name -contains $snapshotname ) {
              $snid = 1 + $snid 
              $snapshotname = "SNAPSHOT#" + $snid + "," + $datestring.ToString() + "," + $env:USERNAME
              }
       $body = New-snapshot -VM $vmobject -Name $snapshotname 
       $subject='Snapshot creation successfull for : '+ $vm         
       } 
 
  ####All commands before this ####
$VerbosePreference = $oldverbose
Send-MailMessage -To $email -Cc "suman.sahu@Invesco.com"  -Subject $subject -From "MBFN6270@invesco.com" -SmtpServer "emailnasmtp.app.invesco.net" -Body $BODY
Cleanup-Variables "End"
}
   else 
    {
        #Write-Host 'Invalid server name given :' $vm +' '+ $email        
        Send-MailMessage -To $email -Cc "suman.sahu@Invesco.com"  -Subject "Snapshort Creation failed : $vm" -From "MBFN6270@invesco.com" -SmtpServer "emailnasmtp.app.invesco.net" -Body 'Invalid server name given'
    }
}



#Clear-Content '\\ushoudsutl01\Scripts\SnapshotCreation\Input_Servers.csv'
#(Get-Content '\\ushoudsutl01\Scripts\SnapshotCreation\Input_Servers.csv' |  Select -First 1) | Out-File '\\ushoudsutl01\Scripts\SnapshotCreation\Input_Servers.csv'
$excel = New-Object -ComObject excel.application
$excel.Visible = $False
$excel.DisplayAlerts = $False
$workbook = $excel.Workbooks.Add()
1..2 | ForEach {
    $Workbook.worksheets.item(2).Delete()
    }
$serverInfoSheet = $workbook.Worksheets.Item(1)
$serverInfoSheet.Activate() | Out-Null
$serverInfoSheet.Cells.Item(1,1)= 'ServerName'
$serverInfoSheet.Cells.Item(1,1).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item(1,1).Font.Bold=$True
$serverInfoSheet.Cells.Item(1,2)= 'SNNumber'
$serverInfoSheet.Cells.Item(1,2).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item(1,2).Font.Bold=$True
$serverInfoSheet.Cells.Item(1,3)= 'Lanid'
$serverInfoSheet.Cells.Item(1,3).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item(1,3).Font.Bold=$True
$serverInfoSheet.Cells.Item(1,4)= 'Requestor'
$serverInfoSheet.Cells.Item(1,4).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item(1,4).Font.Bold=$True
$file = "\\ushoudsutl01\Scripts\SnapshotCreation\Input_Servers.xlsx"
if(Test-Path $file){Remove-Item $file} 
$workbook.SaveAs("$file")
$workbook.Close()





write-host $vm,$desc,$snnumber,$name,$requestor,$vc






