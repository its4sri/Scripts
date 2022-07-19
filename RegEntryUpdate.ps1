$ServerName = Get-Content C:\Servers.csv
foreach ($server in $ServerName) 
{
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v CachedLogonsCount}
if ($CachedLogonsCount = 10)
    {
    Set-ItemProperty -Path $RegKey -Name CachedLogonsCount -Value 60
}