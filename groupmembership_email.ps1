$out = @()

Get-ADGroupMember 'Domain Admins' | ForEach {

    $userDetails = Get-ADUser -Identity $_.SamAccountName

    $props = @{
        SecurityGroup = 'Domain Admins'
        Username = $userDetails.SamAccountName

    }

    $out += New-Object PsObject -Property $props

}

Get-ADGroupMember 'Enterprise Admins' -Server 'houdsvdcwp300.amvescap.net' | ForEach {

    $userDetails = Get-ADUser -Identity $_.SamAccountName

    $props = @{
        SecurityGroup = 'Enterprise Admins'
        Username = $userDetails.SamAccountName

    }

    $out += New-Object PsObject -Property $props

}

Get-ADGroupMember 'Schema Admins' -Server 'houdsvdcwp300.amvescap.net' | ForEach {

    $userDetails = Get-ADUser -Identity $_.SamAccountName

    $props = @{
        SecurityGroup = 'Schema Admins'
        Username = $userDetails.SamAccountName

    }

    $out += New-Object PsObject -Property $props

}


$body  =  "If your d-account is flagged in this report you are responsible to remove it from the respective groups or respond back with a valid justification (Change number or INC details)."
$body  =  $body + "<table border=1><tr><td>GroupName</td><td> UserName </td></tr>"
foreach($obj in $out)
{
    if ($obj.Username -like 'd-*')
    {
   
        $body = $body + "<tr><td style = 'color:red'>"+ $obj.SecurityGroup + "</td><td style = 'color:red'>" +$obj.Username + "</td></tr>"
    }
    else
    {
        $body = $body + "<tr><td style = 'color:green'>"+ $obj.SecurityGroup + "</td><td style = 'color:green'>" +$obj.Username + "</td></tr>"
    }
}

$body = $body + "</table>"
$body

Send-MailMessage -BodyAsHtml -To "Operations.Server@invesco.com" -Cc "IVZ-CyberDefense@invesco.com" , "IVZCORPDomainAdmins@invesco.com" -From "ITInfra-Wintel-DirectoryServices@invesco.com" -Subject PROD:'Enterprise,Domain and Schema admins group membership report' -Body $body  -SmtpServer emailnasmtp.app.invesco.net