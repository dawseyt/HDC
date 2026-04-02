#						Example To Check Remote Access:
$PC = "HD-Teller02"

Get-WmiObject win32_service -ComputerName $pc -Filter "Name='RemoteRegistry' OR Name='RPCSS' OR Name='WinMgmt'" | ForEach-Object { $_.StartService() }

(Get-WmiObject win32_service -ComputerName $PC -Filter "Name='RemoteAccess'").EnableService()

Get-Service -ComputerName $PC -Name RemoteRegistry, RemoteAccess, RpcSs, Winmgmt

##						Start Services For Remote Access:
$PC = "ThatPC"
$services = Get-Service -ComputerName $PC -Name RemoteRegistry, RpcSs, Winmgmt
foreach ($svc in $services) {
    if ($svc.Status -ne "Running") {
        SC \\$PC start $svc.Name
    } else {
        Write-Warning "$($svc.Name) Already running!"
    }
}