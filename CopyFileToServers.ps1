$localpath = "Program Files\Microsoft\Exchange Server\V14\Bin\CmdletExtensionAgents"
$filename = "ScriptingAgentConfig.xml"
$srcfile = "\\USAUSXMR10\E$\$localpath\$filename"
if (Test-Path $srcfile) {
    "Source file: $((Get-Item $srcfile).Length) bytes"
} else {
    "Source path not found: $srcfile" 
    exit
}
Import-Csv D:\scripts\CopyFileToServers\CopyFileToServers-List.csv | foreach {
    $trgpath = "\\$($_.Server)\$($_.Drive)$\$localpath"
    $trgfile = "$trgpath\$filename"
    if (Test-Path $trgpath) {
        if ((Test-Path $trgfile) -and (Get-Item $srcfile).Length -eq (Get-Item $trgfile).Length -and (Get-Item $srcfile).LastWriteTime -eq (Get-Item $trgfile).LastWriteTime) {
            "File already exists on $($_.Server)."
        } else {
            Copy-Item $srcfile $trgpath
            "File copied to $($_.Server)."
            "Target file: $((Get-Item $trgfile).Length) bytes"
        }
    } else {
        "Target path not found: $trgpath" 
    }
    exit
}