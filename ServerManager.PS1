Import-Module ServerManager
Set-ExecutionPolicy -Scope LocalMachine Unrestricted -Force
add-windowsfeature Desktop-Experience,GPMC,IH-Ink-Support,PowerShell-ISE,Remote-Assistance,RDC,RSAT-ADDS-Tools,RSAT-AD-AdminCenter,RSAT-AD-PowerShell,RSAT-RDS-RemoteApp,RSAT-BitLocker,RSAT-Bitlocker-DriveEnc,RSAT-Bitlocker-RecPwd,SNMP-Service,WAS,WAS-Process-Model,WAS-NET-Environment,WAS-Config-APIs,Backup-Features,Backup,Backup-Tools,NET-Framework-Core,Application-Server,FS-Fileserver,FS-Search-Service |Shutdown /r
Set-ExecutionPolicy -Scope LocalMachine restricted -Force
