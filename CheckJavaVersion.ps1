#-----------------------------------------------------------------------
#  This program checks version of Java on each passed Server.
#
#  Created by: Ed Levin
#  Created Date: 01-27-2016
#  Modification History: 
# 
# 
#-----------------------------------------------------------------------

$ListOfServers = Get-Content "D:\scripts\ApplicationSODScripts\CheckJavaVersion\ListOfServers.txt"
$RequiredVersionOfJava = "8.0.40.26"
$JavaLocationPath_1 = "D$\Program Files\Java\jre*\bin\java.exe"
$JavaLocationPath_2 = "D$\Program Files (x86)\Java\jre*\bin\java.exe"
$JavaLocationPath_3 = "C$\Program Files (x86)\Java\jre*\bin\java.exe"

Foreach ($Server in $ListOfServers) {
    $PingStatus = Gwmi Win32_PingStatus -Filter "Address = '$Server'" | Select-Object StatusCode
    IF ($PingStatus.StatusCode -eq 0) {
       IF (Test-Path ("\\$Server\$JavaLocationPath_1")) {
          $Java = gci "\\$Server\$JavaLocationPath_1"
          $Java | Select @{n='Computer';e={$Server}},@{n='JavaVersion';e={if($java.VersionInfo.ProductVersion -eq $RequiredVersionOfJava) {$java.VersionInfo.ProductVersion+'     OK '} ELSE {$java.VersionInfo.ProductVersion+' <-- Not Valid!'} }}
       }
       ELSEIF (Test-Path ("\\$Server\$JavaLocationPath_2")) {
          $Java = gci "\\$Server\$JavaLocationPath_2"
          $Java | Select @{n='Computer';e={$Server}},@{n='JavaVersion';e={$java.VersionInfo.ProductVersion}}
       }
       ELSE {
          $Server | Select @{n='Computer';e={$Server}},@{n='JavaVersion';e={'-'}}
       }
    } #EndIf

    ELSE {
       $Server | Select @{n='Computer';e={$Server}},@{n='JavaVersion';e={'ConnectFailed'}}
    } #EndIf

$smtpServer = "smtp.na.amvescap.com"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "donotreply@invesco.com"
$msg.To.Add("karunakar.maloth@invesco.com)
$msg.Subject = "version of Java"
$msg.Body = ($Server in $ListOfServers) + "For Queries, email <karunakar.maloth@invesco.com>"
$smtp.Send($msg)

} #EndForEach

EXIT 0