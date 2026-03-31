#						Example To Check Remote Access:
$PC = "HD-Teller02"

Get-WmiObject win32_service -ComputerName $pc -Filter "Name='RemoteRegistry' OR Name='RPCSS' OR Name='WinMgmt'" | ForEach-Object { $_.StartService() }

(Get-WmiObject win32_service -ComputerName $PC -Filter "Name='RemoteAccess'").EnableService()

Get-Service -ComputerName $PC -Name RemoteRegistry
Get-Service -ComputerName $PC -Name RemoteAccess
Get-Service -ComputerName $PC -Name RpcSs
Get-Service -ComputerName $PC -Name Winmgmt

##						Start Services For Remote Access:
$PC = "ThatPC"
if ((Get-Service -ComputerName $PC -Name RemoteRegistry).Status -ne "Running") {SC \\$PC start RemoteRegistry} Else {Write-Warning "RemoteRegistry Already running!"}
if ((Get-Service -ComputerName $PC -Name RpcSs).Status -ne "Running") {SC \\$PC start RpcSs} Else {Write-Warning "RPCSS Already running!"}
if ((Get-Service -ComputerName $PC -Name Winmgmt).Status -ne "Running") {SC \\$PC start Winmgmt} Else {Write-Warning "WinMgmt Already running!"}