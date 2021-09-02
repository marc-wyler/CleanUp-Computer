Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

#Clear RecycleBin
Get-ChildItem -Path 'C:\$Recycle.Bin' -Force | Remove-Item -Recurse -Force -Confirm:$false

### Remove all from Public and create ManualBackup
$RemPublic = "C:\Public\"
Get-ChildItem $RemPublic -Recurse | ForEach-Object {Remove-Item $_.FullName -Recurse -Force -Confirm:$false}
$ManualBackup = "C:\Public\ManualBackup" 
If(!(Test-Path $ManualBackup)) {  
    New-Item -Path $ManualBackup -ItemType "Directory"
}

$BackupLog = "D:\source\BackupLog"
If(!(Test-Path $BackupLog)) {  
    New-Item -Path $BackupLog -ItemType "Directory"
}


#Function to export EventLogs
function Backup-Eventlog
{
    param
    (
        [Parameter(Mandatory)]
        [string]
        $LogName,

        [Parameter(Mandatory)]
        [string]
        $DestinationPath
    )

    $eventLog = Get-WmiObject -Class Win32_NTEventLOgFile  -filter "FileName='$LogName'"
    if ($eventLog -eq $null)
    {
        throw "Eventlog '$eventLog' not found."
    }
    
    [int]$status = $eventLog.BackupEventlog($DestinationPath).ReturnValue
    New-Object -TypeName ComponentModel.Win32Exception($status)    
}

#Create and Clear EventLog
Backup-Eventlog -LogName Application -DestinationPath "D:\Source\BackupLog\ApplicationBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx"
Clear-EventLog -LogName Application 
Backup-Eventlog -LogName HardwareEvents -DestinationPath "D:\Source\BackupLog\HardwareBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx" 
Clear-EventLog -LogName HardwareEvents
Backup-Eventlog -LogName "Internet Explorer" -DestinationPath "D:\Source\BackupLog\InternetExplorerBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx" 
Clear-EventLog -LogName "Internet Explorer"
Backup-Eventlog -LogName INtime -DestinationPath "D:\Source\BackupLog\INtimeBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx" 
Clear-EventLog -LogName INtime
Backup-Eventlog -LogName "Key Management Service" -DestinationPath "D:\Source\BackupLog\KeyManagementServiceBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx" 
Clear-EventLog -LogName "Key Management Service"
Backup-Eventlog -LogName Security -DestinationPath "D:\Source\BackupLog\SecurityBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx"
Clear-EventLog -LogName Security 
Backup-Eventlog -LogName System -DestinationPath "D:\Source\BackupLog\SystemBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx"
Clear-EventLog -LogName System
Backup-Eventlog -LogName "Windows PowerShell" -DestinationPath "D:\Source\BackupLog\PowerShellBackup_$((get-date).ToString("yyyy-MM-dd-HHmm")).evtx" 
Clear-EventLog -LogName "Windows PowerShell"


#Windows Error Reportings
$TxtService = "C:\Public\Service.txt"
$Service = Get-Service -Name WerSvc | Out-File $TxtService
$CheckService = Select-String -Path $TxtService -Pattern "Running"
If ($CheckService -ne $null) {
    net stop WerSvc 
    } else {
    Write-Host "Service not running"
}
Remove-Item $TxtService -Recurse -Force -Confirm:$false
$ReportPath = "C:\ProgramData\Microsoft\Windows\WER\ReportQueue"
Get-ChildItem $ReportPath -Include *.wer -Recurse  | ForEach-Object {Remove-Item $_.FullName -Recurse -Force -Confirm:$false} 
Remove-Item $ReportPath -Recurse -Force -Confirm:$false
New-Item -Path "C:\ProgramData\Microsoft\Windows\WER\" -Name "ReportQueue" -ItemType "Directory"

#Check if IP is set to DHCP
$CheckDHCP = "C:\Public\DHCP.txt"
Get-NetIPInterface -InterfaceAlias "LAN1" -AddressFamily IPv4 | Out-File -filepath $CheckDHCP

$CheckFile = Get-Content $CheckDHCP
$containsWord = $CheckFile | ForEach-Object{$_ -match "Enabled"}
if ($containsWord -contains $true) {
    Write-Host "DHCP Enabled"
    [System.Windows.MessageBox]::Show('IP Address is set to DHCP', 'Warning', 'Ok', 'Error')
} else {
    Write-Host "DHCP Disabled"
}

Remove-Item $CheckDHCP -Recurse -Force -Confirm:$false

net start WerSvc #start service
[System.Windows.MessageBox]::Show('Restart the System - check the AS1 and the EventLogs/SystemCheck', 'Restart', 'Ok', 'Information')