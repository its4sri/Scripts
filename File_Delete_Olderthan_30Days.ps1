$now = Get-Date
$Days = "30"
$TargetFolders = Get-Content "D:\scripts\FilePurgeOlderthan30Days\path.txt"
$LastWrite = $now.AddDays(-$Days)
$FIles = foreach ($TargetFolder in $TargetFolders) 
{
Get-ChildItem $TargetFolder | Where {$_.LastWriteTime -le "$lastWrite"} | Remove-Item -Recurse
}
 