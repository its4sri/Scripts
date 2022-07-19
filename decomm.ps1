###########################################
#decomv3.ps1                              #
#Server Decomission script                #
#Run by scheduled task on ushoudsutl01    #
#pulls data from input file decommlist.txt#
#Updated 7/7/2010                         #
###########################################
#
Add-PSSnapin Quest.ActiveRoles.ADManagement
#
# Global Variable Declaration
$datestring = [string](get-date).Month + "." + [string](get-date).day + "." + [string](get-date).year
$logsfolder = "d:\scripts\logs"
set-variable -name MasterLog -value "$logsfolder\$datestring.log"
set-variable -name DecommList -value "d:\scripts\decommlist.txt"
$notapplicablevalue = "not applicable"
# Copy of decommlist for HPSIM removal added by Venkat Dantuluri on 6/12/2015
$HPSIMSourceFile='D:\scripts\decommlist.txt'
#
#Function - writes entry to log file, adds timestamp to each logfile line
Function LogEntry
{param ($entry,$logpath)
  $timestamp = "[$(get-date -uformat %T)]"
  write-output "$timestamp  $entry" >> $logpath
}
#
#Function - uses Ping to check server active state
Function PingCheck
{param ($name)
  $error.clear()
  $ping = New-Object System.Net.NetworkInformation.Ping
  $reply = $ping.send("$name")
  if ($error.count -gt 0){return "error"}
  else {return ($reply.status).ToString()}
}
#
#Function - splits servername and domain name from input FQDN.  Returns error if no domain or invalid domain
Function FQDNVerify
{param ($FQDN)
  $FQDN = $FQDN.replace(" ","")
  if (!$FQDN.contains(".")){return "error"}
  else {
    $servername = ($fqdn.split("."))[0]
    $domainname = $fqdn.substring(((($fqdn.split("."))[0]).length+1))
    return $servername, $domainname
  }
}
#
#Function - returns name of domain controller for input domain
Function GetDC
{param ($domain)
  switch ($domain) {
    "corp.amvescap.net"{return "HOUDSVDCWP100.corp.amvescap.net", "ou=disabled,ou=avz accounts,dc=corp,dc=amvescap,dc=net"}
    "amvescap.net"{return "HOUDSVDCWP300.amvescap.net", $notapplicablevalue}
    "corpdev.dev"{return "usdaldc01d.corpdev.dev", "ou=disabled,ou=accounts,dc=corpdev,dc=dev"}
    "corpuat.uat"{return "USDSVDCWT100.corpuat.uat", "ou=disabled,ou=accounts,dc=corpuat,dc=uat"}
    {"ops.invesco.net" -or "ops.amvescap.net" -or "app.amvescap.net"} {return "ushoudc01.corp.amvescap.net", $notapplicablevalue}
    default{return "error"}
  }
}
#
#Function - removes DNS entries for server in specified domain.  Returns success level after verification
Function RemoveDNS
{param ($name,$domain,$DC)
  if ($domain -eq "ops.invesco.net") {$zone = "invesco.net"}
  else {$zone = $domain}
  $rib1 = "$name" + "-r.corp.amvescap.net"
  $rib2 = "$name" + "ri.corp.amvescap.net"
  $error.clear()
  $DNSRecords = @()
  $serverlog = "$logsfolder\$name.log"
  LogEntry "--------Beginning DNS Removal-----------" $serverlog
  $DNSRecords = $DNSRecords + (get-wmiobject -computername $dc -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_ResourceRecord where OwnerName = '$name.$domain' AND ContainerName = '$zone'")
  $DNSRecords = $DNSRecords + (get-wmiobject -computername HOUDSVDCWP100.corp.amvescap.net -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_ResourceRecord where OwnerName = '$name.rib' AND ContainerName = 'rib'")
  $DNSRecords = $DNSRecords + (Get-WmiObject -computername HOUDSVDCWP100.corp.amvescap.net -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_ResourceRecord where Ownername = '$rib1' AND ContainerName = 'corp.amvescap.net'")
  $DNSRecords = $DNSRecords + (Get-WmiObject -computername HOUDSVDCWP100.corp.amvescap.net -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_ResourceRecord where Ownername = '$rib2' AND ContainerName = 'corp.amvescap.net'")
  #
  $DNSRecords = $DNSRecords + (get-wmiobject -computername $dc -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_PTRType where RecordData = '$name.$domain.' AND ContainerName = '10.in-addr.arpa'")
  $DNSRecords = $DNSRecords + (get-wmiobject -computername HOUDSVDCWP100.corp.amvescap.net -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_PTRType where RecordData = '$name.rib.' AND ContainerName = '10.in-addr.arpa'")
  $DNSRecords = $DNSRecords + (Get-WmiObject -computername HOUDSVDCWP100.corp.amvescap.net -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_PTRType where RecordData = '$rib1.' AND ContainerName = '10.in-addr.arpa'")
  $DNSRecords = $DNSRecords + (Get-WmiObject -computername HOUDSVDCWP100.corp.amvescap.net -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_PTRType where RecordData = '$rib2.' AND ContainerName = '10.in-addr.arpa'")
  if ($error.count -gt 0){
    LogEntry "Error connecting to DNS Server" $serverlog
    return "error"
  }
  foreach ($record in $DNSRecords){
    if ($record -ne $null){
      LogEntry "Deleting Record $($record.TextRepresentation)" $serverlog
      $record.delete()
    }
  }
  foreach ($record in $dnsrecords) {
    if ($record -ne $null){$verify += (get-wmiobject -computername $dc -namespace "root\MicrosoftDNS" -query "select * from MicrosoftDNS_ResourceRecord where OwnerName = '$($record.OwnerName)' AND ContainerName = '$($Record.ContainerName)'")}
  }    
  if ($verify -eq $null){  LogEntry "--------Ending DNS Removal-----------" $serverlog; return "success"}
  else {
    LogEntry "Verification of DNS removal failed.  The following records remain:  " $serverlog
    $verify | %{LogEntry $_.TextRepresentation $serverlog}
    LogEntry "--------Ending DNS Removal-----------" $serverlog
    return "failure"
  }
}
#
#Function - removes WINS entries for server across all WINS servers.  Returns success level after verification
Function RemoveWINS
{param ($name)
  $error.clear()
  $arrWINSservers = @("USDSVDSIPWP100", "USDSVDSIPWP101", "GBDSVDSIPWP100", "INDSVDSIPWP100", "HKDSVDSIPWP100")
  $serverlog = "$logsfolder\$name.log"
  LogEntry "--------Beginning WINS Removal-----------" $serverlog
  #search for server records on each wins server
  foreach ($server in $arrWINSServers){
    LogEntry "Checking server $server" $serverlog
    $query = invoke-expression "netsh wins server \\$server show name name=$name"
    foreach ($record in $query){
      #retrieve relevant data from query output
      if (($record.replace(" ","")).contains("Name:")){
        $data = $record.replace(" ","")
        $hostname=((($data.replace(":","`t")).replace("[","`t")).replace("]","")).split("`t")[1]
        $endchar=(((($data.replace(":","`t")).replace("[","`t")).replace("]","")).split("`t")[2]).replace("h","")
        #delete record and capture the result
        $result = invoke-expression "netsh wins server \\$server delete name name=$hostname endchar=$endchar"
        if ($result | %{$_.contains("does not exist")}){LogEntry "Server $hostname-$endchar successfully removed from $server" $serverlog}
       }
    }
  }
  LogEntry "--------Ending WINS Removal-----------" $serverlog
  #search for server records on each wins server to verify deletion
  foreach ($server in $arrwinsservers){$output = $output + (invoke-expression "netsh wins server \\$server show name name=$name")}
  if ($error.count -gt 0){return "error"}
  if (($output | %{$_.contains($hostname.ToUpper())}) -eq "True"){return "Failure"}
  else {return "success"}
}

#Function - removes McAfee agent from EPO
Function RemoveMcAfee
{param ($name)
  $error.clear()
  $serverlog = "$logsfolder\$name.log"
  [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
  $epoUser = "corp\corp-svc-mcafee"
  $authstr = "FSXLo4IM5g?ixb1:"
   ##################################################################################
  #url is set to Prod environment and running McAfee Removal-----------
  $url="https://ushouepo01v:8443/remote/system.delete?names=$name"
  $wc=new-object System.net.WebClient
  LogEntry "--------Beginning McAfee Removal from PROD-----------" $serverlog
  $wc.Credentials = new-object System.Net.NetworkCredential -ArgumentList ($epoUser,$authstr)
  $Result = $wc.downloadstring("$url")
  if ($Result.Contains("status: -1")){
    
    LogEntry "unable to fetch from Production environment ; hence searching in DEV Environment" $serverlog
	
	#url is being set to DEV Environment and running McAfee Removal-----------
LogEntry "seacrching in Dev"    $serverlog
	$url="https://ushouepo01vd:8443/remote/system.delete?names=$name"
LogEntry "runing in Dev" $serverlog
	LogEntry "--------Beginning McAfee Removal from DEV-----------" $serverlog
    $Result = $wc.downloadstring("$url")
	if ($Result.Contains("status: -1")){
	LogEntry "$error" $serverlog
    return "unable to fetch from both DEV and Production environment "
    }
	else{
	LogEntry "$result" $serverlog
    return "Successfully removed from DEV Environment"
   }
   }
  else {
    LogEntry "$result" $serverlog
    return "Successfully removed from Production Environment"   
   }
   ##########################################################
  LogEntry "--------Ending McAfee Removal-----------" $serverlog
}
#
#Function - Disabled computer account in Active Directory
Function RemoveADObject
{param ($name,$domain,$OU)
  $error.clear()
  $serverlog = "$logsfolder\$name.log"
  LogEntry "--------Beginning AD Removal-----------" $serverlog
  Connect-QADService -service $domain > $null
  $ServerDN = (Get-QADObject -identity $name'$').DN
  if ($error.count -gt 0) {
    LogEntry "Failure searching domain $domain" $serverlog
    Disconnect-QADService
    LogEntry "--------Ending AD Removal-----------" $serverlog
    Return "error"
  }
  if ($ServerDN -eq $null) {
    LogEntry "No computer account found for $name" $serverlog
    Disconnect-QADService
    LogEntry "--------Ending AD Removal-----------" $serverlog
    Return "success"
  }
  if ($OU -eq $notapplicablevalue){
    LogEntry "$name not moved, target OU value is `"$notapplicablevalue`"" $serverlog
    Disconnect-QADService
    LogEntry "--------Ending AD Removal-----------" $serverlog
    Return "success"
  }
  $result = Move-QADObject $ServerDN -newparentcontainer $OU
  LogEntry "$name moved, new location $($result.ParentContainerDN)" $serverlog
  start-sleep -s 300
  #verify location
  $verifyDN = (Get-QADObject -identity $name'$').DN
  LogEntry "--------Ending AD Removal-----------" $serverlog
  Disconnect-QADService
  if (($error.count -eq 0)-and($verifyDN -eq "CN=$name,$OU")){return "success"}
  else {return "failure"}
}  
#
#Function - removes machine from all SCCM servers, returns success level based on verification
#
### Added function to remove machines from SCCM 2012 - D. Maggi
#Function - removes machine from SCCM 2012 servers, returns success level based on verification
Function RemoveSCCM2012
{param ($name)
  $error.clear()
  $serverlog = "$logsfolder\$name.log"
  $siteservername = "usappcmprwp100.corp.amvescap.net"
  $sitecode = "PR1"
  LogEntry "--------Beginning SCCM 2012 Removal-----------" $ServerLog
  #remove computer from primary site server
  LogEntry "Checking site server: $siteservername Site Code: $sitecode" $serverlog
  $systems = get-wmiobject -computername $siteservername -namespace root\SMS\site_$sitecode -class SMS_R_System -filter "name = '$name'"
    if ($error.count -gt 0){
    LogEntry "Error connecting to site server: $siteservername.  Aborting SCCM removal" $serverlog
    LogEntry "--------Ending SCCM 2012 Removal-----------" $ServerLog
    return "error"
    }
    if ($systems -eq $null){
      LogEntry "$name not found in site $sitecode" $serverlog
      continue
    }
    foreach ($system in $systems){
      if ($system.name -eq $name){
        $outname = $system.name
        $outresourceid = $system.resourceid
        LogEntry "Deleting record from SCCM - Name: $outname ID: $outresourceid Server: $siteservername" $serverlog
        $system.Delete()
        if($error.count -gt 0){
          LogEntry "Error deleting record Name: $outname ID: $outresourceid"
          return "failure"
        }
      }
    }
  LogEntry "--------Ending SCCM 2012 Removal-----------" $ServerLog
}
#
### End of added function to remove machines from SCCM 2012 - D. Maggi

### Added function to remove machines from SCOM 2012 - Seshendra Keerty - 09.03.2014
Function RemoveSCOM2012
{param ($name,$fullname)
    $PRDDeleteCollection = $null
    $DEVDeleteCollection = $null
    $error.clear()
    $serverlog = "$logsfolder\$name.log"
    
    ## Load SCOM SDK
    $dummy = [System.Reflection.Assembly]::LoadFrom("D:\SCOMBinaries\Microsoft.EnterpriseManagement.Core.dll")
    $dummy = [System.Reflection.Assembly]::LoadFrom("D:\SCOMBinaries\Microsoft.EnterpriseManagement.OperationsManager.dll")
    $dummy = [System.Reflection.Assembly]::LoadFrom("D:\SCOMBinaries\Microsoft.EnterpriseManagement.OperationsManager.Common.dll")
    $dummy = [System.Reflection.Assembly]::LoadFrom("D:\SCOMBinaries\Microsoft.EnterpriseManagement.Runtime.dll")
    
    ## Function for creating IList Collection
    function New-Collection ( [type] $type ) 
    {
    	$typeAssemblyName = $type.AssemblyQualifiedName;
    	$collection = new-object "System.Collections.ObjectModel.Collection``1[[$typeAssemblyName]]";
    	return ,($collection);
    }
    
    LogEntry "--------Beginning SCOM 2012 Removal-----------" $ServerLog
    ## Declare Management Servers
    $PRDServer = "USHOUSCOMMG02.corp.amvescap.net"
    $DevServer = "USHOUSCOMMG01T.corpuat.uat"
    
    ## Connect to PROD
    $PRDMG = [Microsoft.EnterpriseManagement.ManagementGroup]::Connect($PRDServer)
    If ($error.count -gt 0)
    {
        LogEntry "Error connecting to $PRDServer  Aborting SCOM removal" $serverlog
        LogEntry "--------Ending SCOM 2012 Removal-----------" $ServerLog
        return "Error"
    }
    Else
    {
        LogEntry "Checking object $fullname in PROD." $serverlog
        $PRDAdmin = $PRDMG.GetAdministration()
        $PRDAgents = $PRDAdmin.GetAllAgentManagedComputers()
        Foreach ($PRDAgent in $PRDAgents)
        {
            If ($PRDDeleteCollection -eq $null) 
            {
                $PRDDeleteCollection = New-Collection $PRDAgent.GetType()
            }
            if (@($PRDAgent.PrincipalName -eq $fullname))
            {
	            $PRDDeleteCollection.Add($PRDAgent)
                break
            }
        }
        
        ## Delete Objects from PROD
        If ($PRDDeleteCollection.Count -gt 0) 
        {
            $error.Clear()
            LogEntry "Object $fullname found in PROD. Deleting it." $serverlog
            $PRDDelete = $PRDAdmin.DeleteAgentManagedComputers($PRDDeleteCollection)
            If ($error.count -gt 0)
            {
                LogEntry "Error deleting object $fullname." $ServerLog
                LogEntry "--------Ending SCCM 2012 Removal-----------" $ServerLogs
                return "Failure"
            }
            Else
            {
                LogEntry "--------Ending SCOM 2012 Removal-----------" $ServerLog
                return "Success"   
            }
        }
        Else
        {
            LogEntry "Object $fullname NOT found in PROD." $serverlog
            ## Connect to DEV and loop through Input
            $DEVMG = [Microsoft.EnterpriseManagement.ManagementGroup]::Connect($DEVServer)
            If ($error.count -gt 0)
            {
                LogEntry "Error connecting to $DEVServer  Aborting SCOM removal" $serverlog
                LogEntry "--------Ending SCOM 2012 Removal-----------" $ServerLog
                return "Error"
            }
            Else
            {
                LogEntry "Checking object $fullname in DEV." $serverlog
                $DEVAdmin = $DEVMG.GetAdministration()
                $DEVAgents = $DEVAdmin.GetAllAgentManagedComputers()
                foreach ($DEVAgent in $DEVAgents)
                {
                    if ($DEVDeleteCollection -eq $null) 
                    {
                        $DEVDeleteCollection = new-collection $DEVAgent.GetType()
                    }
                    if (@($DEVAgent.PrincipalName -eq $fullname))
                    {
	                    $DEVDeleteCollection.Add($DEVAgent)
                        break
                    }
                }
    
                ## Delete Objects from DEV
                If ($DEVDeleteCollection.Count -gt 0) 
                {
                    LogEntry "Object $fullname found in DEV. Deleting it." $serverlog
                    $error.Clear()
                    $DEVDelete = $DEVAdmin.DeleteAgentManagedComputers($DEVDeleteCollection)
                    If ($error.count -gt 0)
                    {
                        LogEntry "Error deleting object $fullname." $ServerLog
                        LogEntry "--------Ending SCCM 2012 Removal-----------" $ServerLogs
                        return "Failure"
                    }
                    Else
                    {
                        LogEntry "--------Ending SCOM 2012 Removal-----------" $ServerLog
                        return "Success"      
                    }
                }
                Else
                {
                    LogEntry "Object $fullname NOT found in DEV." $serverlog
                    LogEntry "--------Ending SCOM 2012 Removal-----------" $ServerLog
                    return "Success"    
                }
            }
        }
    }
}

### End of added function to remove machines from SCOM 2012 - Seshendra Keerty - 09.03.2014

# Function to remove decommissioned server from "logon to" list on autosys reboot accounts
function ModifyUser
{param ($name)
  $error.clear()
  Connect-QADService -service 'corp.amvescap.net' > $null
  $user1 = get-qaduser -Identity 'corp-svc-asyswklyrbt' -IncludedProperties 'UserWorkstations'
  $user1newlist = @()
  $user1serverlist = ($user1.UserWorkstations).Split(",")
  foreach ($server in $user1serverlist) {
    if ($server -ne $name) {$user1newlist += $server}
  }
  $user1newproperty = [string]::join(",", $user1newlist)
  set-qaduser -identity $user1.DN -objectattributes @{UserWorkstations=$user1newproperty;} > $null

  $user2 = get-qaduser -Identity 'corp-svc-asyswklyrb2' -IncludedProperties 'UserWorkstations'
  $user2newlist = @()
  $user2serverlist = ($user2.UserWorkstations).Split(",")
  foreach ($server2 in $user2serverlist) {
    if ($server2 -ne $name) {$user2newlist += $server2}
  }
  $user2newproperty = [string]::join(",", $user2newlist)
  set-qaduser -identity $user2.DN -objectattributes @{UserWorkstations=$user2newproperty;} > $null
  Disconnect-QADService
  if ($error.count -gt 0){return "error"}
  else {return "success"}
}
#
#Function - Removes hostname from decomlist
Function RemfromList
{param ($name)
  $oldlist = get-content $decommlist
  $newlist = $oldlist | where-object{$_ -ne $name}
  $newlist | set-content $decommlist
  if($name.length -gt 2){LogEntry "Removing $name from decommlist.txt" $MasterLog}
}
#
#Function - send e-mail on alert or completion
Function ResultMail
{param ($ServerData, $event, $status)
  $serverlog = "$logsfolder\$($ServerData.hostname).log"
  $smtpServer = "emailnasmtp.app.invesco.net"
  $msg = new-object Net.Mail.MailMessage
  $smtp = new-object Net.Mail.SmtpClient($smtpServer)
  $msg.From = "Operations.Server@invesco.com"
  $msg.To.Add("Operations.Server@invesco.com")
  $msg.To.Add("GBL-CSSOpsL1Support@invesco.com")
  # $msg.To.Add("GBLCSSOpsMiddleOffice@amvescap.net")
  #$msg.To.Add("InfrastructureStability-EntMonitoring-Ops@invesco.com")
  $msg.To.Add("software.compliance@invesco.com")
#  $msg.To.Add("Cynthia.Wallace@invesco.com")
#  $msg.To.Add("mike.epperson@invesco.com")
  #customize e-mail based on event
  if ($status -eq "error"){
    $msg.Subject = "Server Decommission Error for $($ServerData.fqdn)"
    $msg.Body = $event
    $smtp.Send($msg)
  }
  if ($status -eq "success"){
    $att = new-object Net.Mail.Attachment($serverlog)
    $msg.Subject = "Decommission complete for $($ServerData.fqdn)"
    $msg.Body = "$event`r`n`r`n"
    $msg.Body+= "Step`t`t`tResult`r`n"
    $msg.Body+= "-----`t`t`t--------`r`n`r`n"
    $msg.Body+= "ADRemoval`t`t$($ServerData.ADRemoval)`r`n"
    $msg.Body+= "DNSRemoval`t`t$($ServerData.DNSRemoval)`r`n"
    $msg.Body+= "WINSRemoval`t`t$($ServerData.WINSRemoval)`r`n"
    $msg.Body+= "McAfeeRemoval`t$($ServerData.McAfeeRemoval)`r`n"
    $msg.Body+= "SCCM2012Removal`t$($ServerData.SCCM2012Removal)`r`n"
    $msg.Body+= "SCOM2012Removal`t$($ServerData.SCOM2012Removal)`r`n"
    $msg.Body+= "LanAdmin`t`t$($ServerData.AsysCheck)`r`n"
    $msg.Attachments.Add($att)
    $smtp.Send($msg)
    $att.Dispose()
  }
}
### Added function to update SIM Servers hosts.txt file based on region - Venkat Dantuluri - added on 6/12/2015
Function RemovalListHPSIM
{param ($HPSIMSourceFile)
$USDestination='\\ushouhpsim01v\NodeRemoval\hosts.txt'
$EUDestination='\\gblonhpsim01v\NodeRemoval\hosts.txt'
$APDestination='\\gblonhpsim01v\NodeRemoval\AP_Removal\hosts.txt'
$NoDestination='D:\scripts\NoSIMDecom.txt'
$FQDNServerName=get-content $HPSIMSourceFile
if($FQDNServerName -eq $null -or $FQDNServerName -eq '')
{
}
else
{
foreach($ServerName in $FQDNServerName)
{
$TrimServerName=$ServerName.Split('.')[0]
$PrdVrtSvr = $TrimServerName.Substring($TrimServerName.Length-1,1)
$DevUatVrtSvr = $TrimServerName.Substring($TrimServerName.Length-2,2)

     if ($PrdVrtSvr -eq "V" -or $DevUatVrtSvr -eq "VD" -or $DevUatVrtSvr -eq "VT")
         {
             Add-Content $NoDestination ($TrimServerName+ " (SIM Decom not required, as it is a Virtual Server)" )
         }
     else
         {
             switch -wildcard ($TrimServerName)
                 {
                     “BA*" {Add-Content $USDestination $TrimServerName }
                     "CA*" {Add-Content $USDestination $TrimServerName }
                     “US*" {Add-Content $USDestination $TrimServerName }
                     "AE*" {Add-Content $EuDestination $TrimServerName }
                     "AT*" {Add-Content $EuDestination $TrimServerName }
                     "BE*" {Add-Content $EuDestination $TrimServerName }
                     "CH*" {Add-Content $EuDestination $TrimServerName }
                     "CZ*" {Add-Content $EuDestination $TrimServerName }
                     “DE*" {Add-Content $EuDestination $TrimServerName }
                     "ES*" {Add-Content $EuDestination $TrimServerName }
                     “FR*" {Add-Content $EuDestination $TrimServerName }
                     "GB*" {Add-Content $EuDestination $TrimServerName }
                     "IE*" {Add-Content $EuDestination $TrimServerName }
                     "IT*" {Add-Content $EuDestination $TrimServerName }
                     "LU*" {Add-Content $EuDestination $TrimServerName }
                     "NL*" {Add-Content $EuDestination $TrimServerName }
                     "SE*" {Add-Content $EuDestination $TrimServerName }
                     "AU*" {Add-Content $APDestination $TrimServerName }
                     "CN*" {Add-Content $APDestination $TrimServerName }
                     "HK*" {Add-Content $APDestination $TrimServerName }
                     "IN*" {Add-Content $APDestination $TrimServerName }
                     "JP*" {Add-Content $APDestination $TrimServerName }
                     "KR*" {Add-Content $APDestination $TrimServerName }
                     "SG*" {Add-Content $APDestination $TrimServerName }
                     "TW*" {Add-Content $APDestination $TrimServerName }
                     default {Add-Content $NoDestination ($TrimServerName+ " (No SIM Decom performed, server name is not in the scope of Invesco approved site code)" ) }
                 }
         }
}
}
}
### End of added function to update SIM Servers hosts.txt file based on region - Venkat Dantuluri - added on 6/12/2015
# 
  
#Main Section - returns from functions populate ServerData Hash Table
#

######## Added function on 6/12/2015 by Venkat Dantluri
RemovalListHPSIM $HPSIMSourceFile
######## End of function on 6/12/2015 by Venkat Dantluri

######## Added 2/14/2013 by Jimmy - making copy of input file to be used by the SCOM decom process.
$decommlistcopy = "d:\scripts\decommlist_SCOM.txt"
Copy-Item $decommlist $decommlistcopy -Force
######## End of 2/14/2013 addition

$serverlist = get-content($decommlist)
foreach ($server in $serverlist) {
  if ($server.length -lt 3){RemfromList $server;continue}
  $ServerData = @{"fqdn" = $server}
  Clear-Variable return
  #call function to split server name and domain name, and error if data is invalid
  $return = fqdnverify($ServerData.fqdn)
  if($return -eq "error"){
    $event = "Error in server list, $($ServerData.fqdn) is not in the correct format.  Removing from decommlist.txt"
    LogEntry $event $MasterLog
    RemfromList $ServerData.fqdn
    ResultMail $ServerData $event $return
    continue
  }
  else{
    $serverData.Add("hostname", $return[0])
    $serverData.Add("domainname", $return[1])
  }
  clear-variable return
  #call function to determine domain controller for domain name, also validate correct domain name
  $return = GetDC($serverData.domainname)
  if ($return -eq "error"){
    $event = "Error in server list, $($ServerData.domainname) is not a supported domain.  Removing from decommlist.txt"
    LogEntry $event $MasterLog
    RemfromList $ServerData.fqdn
    ResultMail $ServerData $event $return
    continue
  }
  else{
    $ServerData.Add("DC", $return[0])
    $ServerData.Add("disableOU", $return[1])
  }
  clear-variable return
  #call function to determine if server is still online
  $return = PingCheck $ServerData.hostname
  $ServerData.Add("PingStatus", $return)
  if ($return -eq "Success"){
    LogEntry "Server $($ServerData.hostname) still online, skipping decomission" $MasterLog
    continue
  }
  elseif ($return -eq "error"){
    $event = "Unable to determine state of $($ServerData.fqdn), skipping decomission"
    LogEntry $event $MasterLog
    RemfromList $ServerData.fqdn
    ResultMail $ServerData $event $return
    continue
  }
  clear-variable return
  #call function to remove DNS entries and return success level after verification
  $return = RemoveDNS $ServerData.hostname $ServerData.domainname $ServerData.DC
  $ServerData.Add("DNSRemoval", $return)
  clear-variable return
  if (($ServerData.domainname -eq "ops.invesco.net") -or ($serverData.domainname -eq "ops.amvescap.net")){
    $ServerData.Add("WINSRemoval", "N/A")
    $ServerData.Add("ADRemoval", "N/A")
    $ServerData.Add("SCCMRemoval", "N/A")
    $ServerData.Add("AsysCheck", "N/A")
    $return = "success"
    $event = "Decommission of $($ServerData.fqdn) complete."
    LogEntry $event $MasterLog
    RemfromList $ServerData.fqdn
    ResultMail $ServerData $event $return
    remove-variable ServerData
    continue
  }  
    
  #call function to remove WINS entries and return success level after verification
  $return = RemoveWINS $ServerData.hostname
  $ServerData.Add("WINSRemoval", $return)
  clear-variable return
  
  #call function to remove McAfee entries and return success level after verification
  $return = RemoveMcAfee $ServerData.hostname
  $ServerData.Add("McAfeeRemoval", $return)
  clear-variable return
  
  #call function to move computer account
  if ($serverData.domainname -ne "amvescap.net"){$return = RemoveADObject $ServerData.hostname $ServerData.domainname $ServerData.disableOU}
  $ServerData.Add("ADRemoval", $return)
  clear-variable return
  
### Added function to remove machines from SCCM 2012 - D. Maggi
  #call function to remove computer from SCCM 2012
  $return = RemoveSCCM2012 $ServerData.hostname
  if ($return -eq $null){$return="success"}
  if ($return -eq "error"){
    Start-Sleep -s 900
    clear-variable return
    $return = RemoveSCCM2012 $ServerData.hostname
    if ($return -eq $null){$return="success"}
  }
  $ServerData.Add("SCCM2012Removal", $return)
  clear-variable return
### End of added function to remove machines from SCCM 2012 - D. Maggi
  
  # call function to remove machine from SCOM 2012
  $return = RemoveSCOM2012 -name $ServerData.hostname -fullname $ServerData.fqdn
  if ($return -eq "error"){
    Start-Sleep -s 900
    clear-variable return
    $return = RemoveSCOM2012 -name $ServerData.hostname -fullname $ServerData.fqdn
  }
  $ServerData.Add("SCOM2012Removal", $return)
  Clear-Variable return
  # End of call function to remove machine from SCOM 2012
  
  #call function to check "logon to" list on autosys accounts
  if ($ServerData.domainname -eq "corp.amvescap.net"){
    $return = ModifyUser $ServerData.hostname
    $ServerData.Add("AsysCheck", $return)
  }
  else {$ServerData.Add("AsysCheck", "N/A")}
  clear-variable return
  #Decom complete!
  $return = "success"
  $event = "Decommission of $($ServerData.fqdn) complete."
  LogEntry $event $MasterLog
  RemfromList $ServerData.fqdn
  ResultMail $ServerData $event $return
  remove-variable ServerData
}

"Processing completed." | Out-File "D:\scripts\ServerDecom\ServerDecomHeartbeat.txt"