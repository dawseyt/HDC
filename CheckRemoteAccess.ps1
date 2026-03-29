#						Example To Check Remote Access:
$PC = "HD-Teller02"

(Get-WmiObject win32_service -ComputerName $pc | Where-Object {$_.Name -eq "RemoteRegistry"}).StartService()
ForEach ($Service in "RemoteRegistry", "RPCSS", "WinMgmt"){(Get-WmiObject win32_service -ComputerName $pc | Where-Object {$_.Name -eq $Service}).StartService()}

(Get-WmiObject win32_service -ComputerName $PC | Where-Object {$_.Name -eq "RemoteAccess"}).EnableService()

Get-Service -ComputerName $PC -Name RemoteRegistry
Get-Service -ComputerName $PC -Name RemoteAccess
Get-Service -ComputerName $PC -Name RpcSs
Get-Service -ComputerName $PC -Name Winmgmt

##						Start Services For Remote Access:
$PC = "ThatPC"
if ((Get-Service -ComputerName $PC -Name RemoteRegistry).Status -ne "Running") {SC \\$PC start RemoteRegistry} Else {Write-Warning "RemoteRegistry Already running!"}
if ((Get-Service -ComputerName $PC -Name RpcSs).Status -ne "Running") {SC \\$PC start RpcSs} Else {Write-Warning "RPCSS Already running!"}
if ((Get-Service -ComputerName $PC -Name Winmgmt).Status -ne "Running") {SC \\$PC start Winmgmt} Else {Write-Warning "WinMgmt Already running!"}