powershell.exe -c ". \"C:\Program Files (x86)\VMware\Infrastructure\PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1\" $true; C:\scripts\script.ps1"

$Vservers = @("usdalvc01v","ushouvc01v","gbwokvc01v","gblonvc01v","inbcpvc01v","inhydvc01v")

$gdat = get-date
$fm =$gdat.tostring('MMM_dd_yyyy_hhmmss')

foreach ($sr in $Vservers) {
 
Connect-vIserver -Server $sr
Get-VMHOST -state maintenance | select @{Name="VcentreServer";expression={$sr}},name,connectionstate | export-csv c:\temp\$($fm).csv -Append -NoTypeInformation
disConnect-vIserver -Server $sr -confirm:$false
  
 }
$file="C:\temp\"+$fm+".csv"
#$alarms | Export-CSV "c:\temp\main.csv"
$smtpServer = "smtp.na.amvescap.com"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "donotreply@invesco.com"
$msg.To.Add("ashokkumar.chukka@invesco.com")
#$msg.To.Add("hemamallikarjuna.yakkala@invesco.com")
$msg.Subject = "VMware ESXi Hosts in Maintenance Mode $date"
$msg.IsBodyHTML = $true
$att = new-object Net.Mail.Attachment($file)
$msg.Attachments.Add($att)
$smtp.Send($msg)

Exit


