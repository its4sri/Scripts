$hostname = "USHOURDMAGR01"
$smtpServer = 'emailnasmtp.app.invesco.net' 
$from = "operations.server@invesco.com" 
$recipients = 'hemamallikarjuna.yakkala@invesco.com' 
$Subject = "Restarting Snap Manager Service on $hostname" 
$body = "This is an automated message to confirm that the Snap Manager Service on $hostname  
is about to be restarted as part of agreed scheduled time 3 AM IST every day, please ignore any alerts  
for the next 10 minutes from this service." 
  
  
Send-MailMessage -To $recipients -Subject $Subject -Body $body -From $from -SmtpServer $smtpServer 
  
  
# Stop service 
  
$service = 'SnapManager Service' 
  
Net STOP $service -Verbose 
  
    do { 
        Start-sleep -s 5 
        }  
            until ((get-service $service).Status -eq 'Stopped') 
  
  
  
# Start service 
  
        Net Start $service -Verbose 
  
    do { 
        Start-sleep -s 5 
        }  
            until ((get-service $service).Status -eq 'Running') 
  
  
  
# Send confirmation that service has restarted successfully 
  
$Subject = "Snap Manager Service Restarted Successfully on $hostname" 
$body = "This mail confirms that the Snap Manager service on $hostname is now running. 
Please check your processes status and reach out to Server Operations Team in case of any issues with Snap Manager Service" 
  
Start-Sleep -s 5 
  
Send-MailMessage -To $recipients -Subject $Subject -Body $body -From $from -SmtpServer $smtpServer