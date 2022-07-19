$vCenters=@("ushouvc01v","usdalvc01v","GBLONVC01V","GBWOKVC01V","inhydvc01v","inbcpvc01v")
$gdat = get-date
$fm =$gdat.tostring('MMM_dd_yyyy_hh')
$file="C:\temp\Alaram_Dashboard_$($fm)hrs.csv"
Function Get-TriggeredAlarms {

$rootFolder = Get-Folder -Server $vc "Datacenters"
 
foreach ($vca in $rootFolder.ExtensionData.TriggeredAlarmState) {
$alarm = "" | Select-Object VC, EntityType, Alarm, Entity, Status, Time, Acknowledged, AckBy, AckTime
$alarm.VC = $vCenter
$alarm.Alarm = (Get-View -Server $vc $vca.Alarm).Info.Name
$entity = Get-View -Server $vc $vca.Entity
$alarm.Entity = (Get-View -Server $vc $vca.Entity).Name
$alarm.EntityType = (Get-View -Server $vc $vca.Entity).GetType().Name 
$alarm.Status = $vca.OverallStatus
$alarm.Time = $vca.Time
$alarm.Acknowledged = $vca.Acknowledged
$alarm.AckBy = $vca.AcknowledgedByUser
$alarm.AckTime = $vca.AcknowledgedTime 
$alarm
}
Disconnect-VIServer $vCenter -Confirm:$false
}
 
Write-Host ("Getting the alarms from {0} vCenters." -f $vCenters.Length)
 
$alarms = @()
foreach ($vCenter in $vCenters) {
        $vc = Connect-VIServer $vCenter
Write-Host "Getting alarms from $vCenter."
$alarms += Get-TriggeredAlarms $vCenter
}
 
$alarms |export-csv $file -Append -NoTypeInformation

$file = "C:\temp\Alaram_Dashboard_$($fm)hrs.csv"
$alarms | Export-CSV "C:\temp\Alaram_Dashboard_$($fm)hrs.csv"
$smtpServer = "smtp.na.amvescap.com"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "donotreply@invesco.com"
$msg.To.Add("GBLCSSOpsPlatform@invesco.com")
#$msg.To.Add("hemamallikarjuna.yakkala@invesco.com")
$msg.Subject = "vCenter Alarms $fm"
$msg.IsBodyHTML = $true
$att = new-object Net.Mail.Attachment($file)
$msg.Attachments.Add($att)
$smtp.Send($msg)