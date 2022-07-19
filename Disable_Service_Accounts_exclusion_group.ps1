############################################################################################################################
# AUTHOR       : Jonathan Turner
# DATE         : Date: May 15, 2017
# EDIT         :
# COMMENT      :This script will check the lastLogonTimestamp attribute on service account objects.  If this date is older than 180 days
#              :the computer object will be disabled and moved to the disabled OU for each domain.
# VERSION      :1.0
############################################################################################################################

#------------------------------------------------------------------------------------------------------------------------------------
# CREATING LOG FILE
#------------------------------------------------------------------------------------------------------------------------------------

#Gathering Data for Log File Creation#
$day = (Get-Date).Day.ToString()
$month = (Get-Date).Month.ToString()
$year = (Get-Date).Year.ToString()
$moveCount = 0
$deleteCount = 0
$logDate = $day + "_" + $month + "_" +$year
$logPath = "D:\Scripts\Logs\Inactive_User_Accounts\Inactive_Users_" + $logDate + ".log"
$detailedLogPath = "D:\Scripts\Logs\Inactive_User_Accounts\Detailed_Inactive_Users_" + $logDate + ".log"
$Members = Get-ADGroupMember -Identity 'CORP-G-POL-InactiveAccountException'
$Heartbeatpath = "D:\Scripts\Logs\Inactive_User_Accounts"

#Checking if log file exists.  Log file created if it does not exist and cleared if it does exist
if(!(test-path -Path $logPath))
{
    New-Item -Path $logPath -ItemType File
}
else
{
    Clear-Content $logPath
}

if(!(test-path -Path $detailedLogPath))
{
    New-Item -Path $detailedLogPath -ItemType File
}
else
{
    Clear-Content $detailedLogPath
}


$testDate= (Get-Date).AddDays(-180) #Today's date minus 180 days
Add-Content -Path $logPath -Value "************************************************************************************************************************"
Add-Content -Path $detailedLogPath -Value "************************************************************************************************************************"
Add-Content -Path $logPath -Value "[$([DateTime]::Now)]:  Script started."
Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]:  Script started."
Add-Content -Path $logPath -Value "[$([DateTime]::Now)]:  Disabling accounts that have not loggedin for more than 180 days  $testDate"
Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]:  Disabling accounts that have not loggedin for more than 180 days  $testDate"

#------------------------------------------------------------------------------------------------------------------------------------
# GENERATING COMPUTER LIST
#------------------------------------------------------------------------------------------------------------------------------------

#Determining current domain
$currentDomain = (Get-ADDomain).DistinguishedName
$ldapPath = "OU=Disabled,OU=AVZ Accounts," + $currentDomain
Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]:  Searching $currentDomain for Windows computers."
$moveCount = 0
#Generating Computer List for the domain

$userList = @()
$searchOU = @()
$searchOU += "OU=Application,OU=AVZ Accounts,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Automation,OU=AVZ Accounts,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Service,OU=AVZ Accounts,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Service Accounts,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Training,OU=AVZ Accounts,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Service Accounts,OU=UNIX,OU=AVZ Special Purpose,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Unused,OU=AVZ Accounts,DC=corp,DC=amvescap,DC=net"
$searchOU += "OU=Service Accounts,OU=UNIX,DC=corp,DC=amvescap,DC=net"

foreach($ou in $searchOU)
{
#$ou = "OU=Unused,OU=AVZ Accounts,DC=corpdev,DC=dev"
    Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]: Searching $ou in found in the domain $currentDomain."
    
    $ouAccounts = Search-ADAccount -AccountInactive -TimeSpan([timespan]180d) -SearchBase $ou -SearchScope Subtree 
    $userCount = $ouAccounts.Count
    
    Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]: $userCount user accounts found in the OU $ou."
    foreach ($account in $ouAccounts)
    {
        $isMemeberexists = $Members | ?{ $_.SamAccountName -eq $account.SamAccountName}
        if(!$isMemeberexists){
            $shortName = $account.SamAccountName
            $userCreatedDate = get-aduser -Identity $shortName -Properties * |select whenCreated
            $userdate = $userCreatedDate.whenCreated.ToString("yyyy-MM-dd")
            $currentDate = (Get-Date).ToString("yyyy-MM-dd")
            $diff =(New-TimeSpan -Start  $userdate -end  $currentDate).Days  
            if($diff -gt 180)
            {
                Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]: The account $shortName has not logged in within the last 180 days and moved to the disabled OU."
                Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]: The account $shortName has been moved to the disabled OU."
                Add-Content -Path $detailedLogPath -Value "$shortName"
                Set-ADUser $account -Enabled $false
                Write-Host "Disabling account $shortName"
                Move-ADObject -identity $account -TargetPath $ldapPath
                $moveCount = $moveCount + 1
            }
            else
            {
                 Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]: The account $shortName is created less than 180 days"
            }

            
        }
        Else{
            Add-Content -Path $detailedLogPath -Value "[$([DateTime]::Now)]: The account $($isMemeberexists.SamAccountName) is in exception list."
        }
    }
}
"Processing completed." | Out-File "$Heartbeatpath\Disable_Inactiveaccounts_Heartbeat.txt"