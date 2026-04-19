# ============================================================================
# RemoteManagement.psm1 - Remote System and Session Management
# ============================================================================

function Get-RemoteProcesses {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [switch]$AsJob
    )
    
    $cmdArgs = @{
        ComputerName = $ComputerName
        ScriptBlock = {
            Get-Process | Select-Object Name, Id, 
                @{Name='CPU';Expression={if($_.CPU){[math]::Round($_.CPU, 2)}else{0}}}, 
                @{Name='MemMB';Expression={if($_.WorkingSet64){[math]::Round($_.WorkingSet64 / 1MB, 2)}else{0}}}, 
                Description -ErrorAction SilentlyContinue
        }
        ErrorAction = 'Stop'
    }
    
    if ($AsJob) {
        $cmdArgs.Add('AsJob', $true)
    }
    
    Invoke-Command @cmdArgs
}

function Stop-RemoteProcess {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [int]$ProcessId
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
        param($pidToKill) 
        Stop-Process -Id $pidToKill -Force -ErrorAction Stop
    } -ArgumentList $ProcessId -ErrorAction Stop
}

function Get-RemoteActiveUsers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    $quserOutput = quser /server:$ComputerName 2>&1
    
    if ($LASTEXITCODE -ne 0 -or $quserOutput -match "No User exists" -or $quserOutput -match "Error") {
        return @()
    }
    
    $sessions = @()
    for ($i = 1; $i -lt $quserOutput.Count; $i++) {
        $line = $quserOutput[$i] -replace '^>', ' ' # Remove active session indicator
        
        $uName = $null
        $sId = $null
        $sState = $null
        
        if ($line.Length -ge 65) {
            # Fixed-width parsing for reliability
            $uName = $line.Substring(0, 22).Trim()
            $sId = $line.Substring(41, 5).Trim()
            $sState = $line.Substring(46, 8).Trim()
        } else {
            # Fallback regex/split parsing
            $tok = $line -split '\s+' | Where-Object { $_ }
            if ($tok.Count -ge 5) {
                $uName = $tok[0]
                $sId = ($tok | Where-Object { $_ -match '^\d+$' } | Select-Object -First 1)
                $sState = if ($line -match 'Active') { 'Active' } else { 'Disc' }
            }
        }
        
        if ($uName -and $sId) {
            $sessions += [PSCustomObject]@{
                Username  = $uName
                SessionId = $sId
                State     = $sState
            }
        }
    }
    return $sessions
}

function Stop-RemoteUserSession {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$SessionId
    )
    
    $logoffRes = logoff $SessionId /server:$ComputerName 2>&1
    
    if ($LASTEXITCODE -ne 0 -or (-not [string]::IsNullOrWhiteSpace($logoffRes) -and $logoffRes -match "Error")) {
        throw "Failed to logoff session $SessionId on ${ComputerName}. Details: $logoffRes"
    }
    return $true
}

function Invoke-RemoteGPResult {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFile,
        
        [string]$TargetUser = $null
    )
    
    $gpArgs = @("/S", $ComputerName)
    if (-not [string]::IsNullOrWhiteSpace($TargetUser)) {
        $gpArgs += "/USER", $TargetUser
    } else {
        $gpArgs += "/SCOPE", "COMPUTER"
    }
    $gpArgs += "/H", $OutputFile, "/F"
    
    $p = Start-Process -FilePath "gpresult.exe" -ArgumentList $gpArgs -NoNewWindow -PassThru -Wait
    
    if ($p.ExitCode -ne 0 -or -not (Test-Path $OutputFile)) {
        throw "GPResult generation failed. Process returned exit code $($p.ExitCode)."
    }
    
    return [PSCustomObject]@{
        Success    = $true
        OutputFile = $OutputFile
        ExitCode   = $p.ExitCode
    }
}

function Get-RemotePrinterDrivers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
        Get-PrinterDriver | Select-Object -ExpandProperty Name 
    } -ErrorAction Stop
}

function Get-RemotePrinters {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
        Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published, DeviceType, PrinterStatus, Location, Comment
    } -ErrorAction Stop
}

function Remove-RemotePrinter {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$PrinterName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($pName)
        Remove-Printer -Name $pName -ErrorAction Stop
    } -ArgumentList $PrinterName -ErrorAction Stop
}

function Install-RemotePrinter {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$PrinterName,
        
        [Parameter(Mandatory=$true)]
        [string]$DriverName,
        
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )
    
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($PName, $PDriver, $PIP)
        
        $portName = "IP_$PIP"
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $PIP -ErrorAction Stop
        }
        
        if (-not (Get-PrinterDriver -Name $PDriver -ErrorAction SilentlyContinue)) {
            Add-PrinterDriver -Name $PDriver -ErrorAction Stop
        }

        Add-Printer -Name $PName -DriverName $PDriver -PortName $portName -ErrorAction Stop
        return $true
    } -ArgumentList $PrinterName, $DriverName, $IPAddress -ErrorAction Stop
}

function Deploy-DriverToRemote {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourceInfPath,
        
        [Parameter(Mandatory=$true)]
        [string]$RemoteComputer
    )
    
    $parentDir = [System.IO.Path]::GetDirectoryName($SourceInfPath)
    $folderName = [System.IO.Path]::GetFileName($parentDir)
    $infName = [System.IO.Path]::GetFileName($SourceInfPath)
    
    $destPath = "\\$RemoteComputer\C`$\Temp\Drivers\Upload_$(Get-Date -Format 'yyyyMMddHHmmss')_$folderName"
    
    if (-not (Test-Path $destPath)) { 
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null 
    }
    
    # Copy all files from the driver directory to remote temp location
    Copy-Item -Path "$parentDir\*" -Destination $destPath -Recurse -Force -ErrorAction Stop
    
    $localDest = $destPath.Replace("\\$RemoteComputer\C`$", "C:")
    
    # Execute PnPUtil on remote computer to stage the driver
    $pnpOutput = Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
        param($Path, $InfFile)
        $driverPath = Join-Path $Path $InfFile
        if (Test-Path $driverPath) {
            return (pnputil.exe /add-driver "$driverPath" /install)
        } else {
            throw "Driver file not found on remote machine at $driverPath"
        }
    } -ArgumentList $localDest, $infName -ErrorAction Stop
    
    return $pnpOutput
}

function Get-RemoteUptime {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    $job = Start-Job -ScriptBlock {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $args[0] -ErrorAction Stop
        $lastBoot = $os.LastBootUpTime
        $uptime = (Get-Date) - $lastBoot
        return [PSCustomObject]@{ ComputerName = $args[0]; LastBootUpTime = $lastBoot; Days = $uptime.Days; Hours = $uptime.Hours; Minutes = $uptime.Minutes }
    } -ArgumentList $ComputerName
    
    $tCount = 100
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Connection timed out after 10 seconds."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Restart-RemoteSpooler {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Restart-Service "Spooler" -Force -ErrorAction Stop
    } -ErrorAction Stop
}

function Invoke-RemotePowerAction {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Restart", "Shutdown")]
        [string]$Action
    )
    if ($Action -eq "Restart") {
        Restart-Computer -ComputerName $ComputerName -Force -ErrorAction Stop
    } else {
        Stop-Computer -ComputerName $ComputerName -Force -ErrorAction Stop
    }
}

function Start-RemoteProcess {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$CommandLine
    )
    $job = Start-Job -ScriptBlock {
        param($c, $cmd)
        $proc = Invoke-CimMethod -ClassName Win32_Process -ComputerName $c -MethodName Create -Arguments @{ CommandLine = $cmd } -ErrorAction Stop
        if ($proc.ReturnValue -ne 0) {
            throw "WMI Return Code: $($proc.ReturnValue). (2 = Access Denied, 3 = Insufficient Privilege, 9 = Path Not Found)"
        }
    } -ArgumentList $ComputerName, $CommandLine
    
    $tCount = 100
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Connection timed out (10s). The computer may be offline or firewalled."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Get-RemoteServices {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    $job = Start-Job -ScriptBlock {
        Get-CimInstance -ClassName Win32_Service -ComputerName $args[0] -ErrorAction Stop | Select-Object Name, DisplayName, State, StartMode
    } -ArgumentList $ComputerName
    
    $tCount = 100
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Connection timed out (10s). The computer may be offline or firewalled."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Invoke-RemoteServiceAction {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Start", "Stop", "Restart")]
        [string]$Action
    )
    
    $job = Start-Job -ScriptBlock {
        param($c, $s, $a)
        $svc = Get-CimInstance -ClassName Win32_Service -ComputerName $c -Filter "Name='$s'" -ErrorAction Stop
        if (-not $svc) { throw "Service '$s' not found on target." }

        if ($a -eq "Start") {
            Invoke-CimMethod -InputObject $svc -MethodName StartService -ErrorAction Stop | Out-Null
        } elseif ($a -eq "Stop") {
            Invoke-CimMethod -InputObject $svc -MethodName StopService -ErrorAction Stop | Out-Null
        } elseif ($a -eq "Restart") {
            Invoke-CimMethod -InputObject $svc -MethodName StopService -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Seconds 2
            Invoke-CimMethod -InputObject $svc -MethodName StartService -ErrorAction Stop | Out-Null
        }
    } -ArgumentList $ComputerName, $ServiceName, $Action
    
    $tCount = 150
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Command timed out (15s)."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Get-RemoteDiskSpace {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop | Select-Object Size, FreeSpace
    } -ErrorAction Stop
}

function Get-RemoteUserProfiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance -ClassName Win32_UserProfile -Filter "Special=False" -ErrorAction Stop | Select-Object LocalPath, LastUseTime, Loaded, SID
    } -ErrorAction Stop
}

function Remove-RemoteUserProfile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$SID
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($s)
        $prof = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$s'" -ErrorAction Stop
        Invoke-CimMethod -InputObject $prof -MethodName Delete -ErrorAction Stop
    } -ArgumentList $SID -ErrorAction Stop
}

# ----------------------------------------------------------------------------
# NOTE: The below functions support the existing System Manager (ProcessManager) tab.
# ----------------------------------------------------------------------------
function Get-RemoteInstalledSoftware {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        $paths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
        try {
            Get-ItemProperty $paths -ErrorAction Stop | 
                Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 -and $_.ParentKeyName -eq $null } | 
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, UninstallString, QuietUninstallString | 
                Sort-Object DisplayName -Unique
        } catch { }
    } -ErrorAction SilentlyContinue
}

function Repair-HDRemoteSoftware {
    param (
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][string]$Identifier,
        [Parameter(Mandatory=$true)][string]$Type
    )

    $scriptBlock = {
        param($targetId, $targetType)
        try {
            if ($targetType -eq 'AppX') {
                # 🛡️ Sentinel: Base64 encode untrusted input to prevent command injection
                $encodedId = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($targetId))
                # For AppX, a repair is often essentially re-registering the manifest
                $psCmd = "`$id = [System.Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedId')); Get-AppxPackage -Name `"*`$id*`" -AllUsers | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register `"`$($_.InstallLocation)\AppXManifest.xml`"}"
                $encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($psCmd))
                Start-Process powershell.exe -ArgumentList @("-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", $encodedCmd) -Wait -WindowStyle Hidden
                return $true
            } else {
                # Look for MSI product codes which are enclosed in {}
                if ($targetId -match '\{[a-fA-F0-9-]{36}\}') {
                    $productCode = $matches[0]
                    # fa = force repair all
                    $cmd = "msiexec.exe /fa $productCode /qn /norestart"
                    Start-Process cmd.exe -ArgumentList @("/c", $cmd) -Wait -WindowStyle Hidden
                    return $true
                } else {
                    throw "Automatic repair is only supported for MSI-based installers (needs a Product Code GUID)."
                }
            }
        } catch { throw $_ }
    }

    try {
        Invoke-Command -ComputerName $ComputerName -ArgumentList $Identifier, $Type -ScriptBlock $scriptBlock -ErrorAction Stop
        return $true
    } catch {
        Write-Warning "Failed to repair $Identifier on $($ComputerName): $($_.Exception.Message)"
        return $false
    }
}

function Uninstall-RemoteSoftware {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        [string]$QuietUninstallString,
        [string]$UninstallString
    )
    $cmd = $QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($cmd)) {
        $cmd = $UninstallString
        if ([string]::IsNullOrWhiteSpace($cmd)) { throw "No uninstall string found for this application." }

        $exe = ''
        $argsString = ''

        if ($cmd -match '^\s*"([^"]+)"(.*)$') {
            $exe = $matches[1]
            $argsString = $matches[2].Trim()
        } else {
            $possibleExe = ''
            $remainingArgs = ''
            $parts = $cmd -split '\s+'

            for ($i = 0; $i -lt $parts.Count; $i++) {
                if ($i -gt 0) { $possibleExe += ' ' }
                $possibleExe += $parts[$i]

                if ($possibleExe -match '(?i)\.exe$') {
                    $exe = $possibleExe
                    if ($i + 1 -lt $parts.Count) {
                        $remainingArgs = ($parts[($i+1)..($parts.Count-1)]) -join ' '
                    }
                    break
                }
            }

            if (-not $exe) {
                $parts = $cmd -split '\s+', 2
                $exe = $parts[0]
                if ($parts.Count -gt 1) {
                    $remainingArgs = $parts[1]
                }
            }
            $argsString = $remainingArgs
        }

        $argsArray = @()
        if (-not [string]::IsNullOrWhiteSpace($argsString)) {
            $regex = '(?:[^\s"]|"[^"]*")+'
            $matchesArg = [regex]::Matches($argsString, $regex)
            foreach ($m in $matchesArg) {
                $argsArray += $m.Value
            }
        }

        if ($exe -match '(?i)msiexec(\.exe)?$') {
            for ($i = 0; $i -lt $argsArray.Count; $i++) {
                $argsArray[$i] = $argsArray[$i] -replace '(?i)/I', '/X'
            }
            $joinedArgs = $argsArray -join ' '
            if ($joinedArgs -notmatch '(?i)/q') {
                $argsArray += '/qn'
                $argsArray += '/norestart'
            }
        } elseif ($exe -match '(?i)unins\d{3}\.exe$') {
            $joinedArgs = $argsArray -join ' '
            if ($joinedArgs -notmatch '(?i)/VERYSILENT') { $argsArray += '/VERYSILENT' }
            if ($joinedArgs -notmatch '(?i)/SUPPRESSMSGBOXES') { $argsArray += '/SUPPRESSMSGBOXES' }
            if ($joinedArgs -notmatch '(?i)/NORESTART') { $argsArray += '/NORESTART' }
        } elseif ($exe -match '(?i)uninstall\.exe$') {
            $joinedArgs = $argsArray -join ' '
            if ($joinedArgs -notmatch '(?i)/S') { $argsArray += '/S' }
        }

        $cmd = "`"$exe`" $($argsArray -join ' ')"
    }
    Start-RemoteProcess -ComputerName $ComputerName -CommandLine $cmd
}

function Get-RemoteDevices {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-PnpDevice -ErrorAction SilentlyContinue | Select-Object FriendlyName, Class, Status, Manufacturer, InstanceId | Where-Object { -not [string]::IsNullOrWhiteSpace($_.FriendlyName) }
    } -ErrorAction Stop
}

function Set-RemoteDeviceState {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$InstanceId,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Enable", "Disable")]
        [string]$Action
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($id, $act)
        if ($act -eq 'Enable') {
            Enable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        } else {
            Disable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        }
    } -ArgumentList $InstanceId, $Action -ErrorAction Stop
}

function Get-RemoteEventLogs {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        try {
            Get-WinEvent -FilterHashtable @{LogName='System','Application'; Level=1,2,3} -MaxEvents 150 -ErrorAction Stop |
                Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message
        } catch { @() }
    } -ErrorAction Stop
}

function Get-RemoteDeviceInfo {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    try {
        $cleanFnString = ${function:Clean-WmiString}.ToString()
        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            param($cleanFnStringLocal)
            # Define helper locally within the scriptblock
            ${function:Clean-WmiString} = [scriptblock]::Create($cleanFnStringLocal)

            $cs  = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $bio = Get-CimInstance Win32_BIOS            -ErrorAction Stop
            $bat = Get-CimInstance Win32_Battery          -ErrorAction SilentlyContinue

            $model = if ($cs.Manufacturer -match "LENOVO" -and $cs.Model.Length -ge 4) {
                $cs.Model.Substring(0, 4)
            } else { $cs.Model }

            $batteryStatus = if ($bat) { "$($bat.EstimatedChargeRemaining)%" } else { "No Battery / Desktop" }

            # Map the integer status to a human-readable string
            $adminPwdStatus = switch ($cs.AdminPasswordStatus) {
                0 { "Disabled" }
                1 { "Enabled" }
                2 { "Not Implemented" }
                3 { "Unknown" }
                default { if ($null -ne $cs.AdminPasswordStatus) { "Unknown ($($cs.AdminPasswordStatus))" } else { "Unknown" } }
            }

            $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
            $freeGB = if ($disk) { [math]::Round($disk.FreeSpace / 1GB, 1) } else { 0 }
            $sizeGB = if ($disk) { [math]::Round($disk.Size / 1GB, 1) } else { 0 }
            $freePct = if ($sizeGB -gt 0) { [math]::Round(($freeGB / $sizeGB) * 100, 1) } else { 0 }

            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            $uptimeDays = if ($os) { ([math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)) } else { 0 }

            $pendingReboot = $false
            try {
                if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pendingReboot = $true }
                if (-not $pendingReboot) {
                    $sm = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
                    if ($null -ne $sm) { $pendingReboot = $true }
                }
            } catch {}

            return [PSCustomObject]@{
                ComputerName        = Clean-WmiString $cs.Name
                Manufacturer        = Clean-WmiString $cs.Manufacturer
                Model               = Clean-WmiString $model
                SystemFamily        = Clean-WmiString $cs.SystemFamily
                SerialNumber        = Clean-WmiString $bio.SerialNumber
                BIOSVersion         = Clean-WmiString $bio.SMBIOSBIOSVersion
                BatteryStatus       = Clean-WmiString $batteryStatus
                AdminPasswordStatus = $adminPwdStatus
                QueryTime           = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                CDriveFreePct       = $freePct
                UptimeDays          = $uptimeDays
                PendingReboot       = $pendingReboot
            }
        } -ArgumentList $cleanFnString -ErrorAction Stop

        return [PSCustomObject]@{
            ComputerName        = $result.ComputerName
            Manufacturer        = $result.Manufacturer
            Model               = $result.Model
            SystemFamily        = $result.SystemFamily
            SerialNumber        = $result.SerialNumber
            BIOSVersion         = $result.BIOSVersion
            BatteryStatus       = $result.BatteryStatus
            AdminPasswordStatus = $result.AdminPasswordStatus
            QueryTime           = $result.QueryTime
            CDriveFreePct       = $result.CDriveFreePct
            UptimeDays          = $result.UptimeDays
            PendingReboot       = $result.PendingReboot
        }
    } catch {
        Write-Warning "Get-RemoteDeviceInfo failed for $ComputerName`: $($_.Exception.Message)"
        return $null
    }
}

function Get-RemoteScreenshot {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    $bytes = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop

        $screen   = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bitmap   = [System.Drawing.Bitmap]::new($screen.Width, $screen.Height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
        $graphics.Dispose()

        $stream = [System.IO.MemoryStream]::new()
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        $bitmap.Dispose()
        return $stream.ToArray()
    } -ErrorAction Stop

    return $bytes
}

function Test-RemotePrinterPort {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PrinterIP
    )
    $result = [PSCustomObject]@{
        Online   = $false
        Ping     = $false
        Port9100 = $false
    }

    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send($PrinterIP, 1500)
        if ($reply.Status -eq 'Success') { $result.Ping = $true; $result.Online = $true }
    } catch {}

    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($PrinterIP, 9100, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne(1500, $false)
        if ($ok -and $tcp.Connected) { $result.Port9100 = $true; $result.Online = $true }
        $tcp.Close()
    } catch {}

    return $result
}

function Get-RemotePrintJobs {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-PrintJob -PrinterName * -ErrorAction SilentlyContinue |
            Select-Object Id, PrinterName, DocumentName, UserName,
                          TotalPages, PagesPrinted, JobStatus,
                          @{ Name='SizeKB'; Expression={ [math]::Round($_.Size / 1KB, 1) } },
                          SubmittedTime
    } -ErrorAction Stop
}

# ----------------------------------------------------------------------------
# NEW ENHANCEMENTS: Dedicated AppX Software Manager & Wake-On-LAN
# ----------------------------------------------------------------------------

<#
.SYNOPSIS
    Retrieves installed software and AppX packages by querying the registry
    directly to bypass WinRM WinRT loading crashes.
#>
function Get-HDRemoteSoftware {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )

    $scriptBlock = {
        $out = @()
        
        # 1. Standard Desktop Apps (Registry)
        $registryPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        foreach ($path in $registryPaths) {
            try {
                $items = Get-ItemProperty $path -ErrorAction Stop | Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 -and $_.ParentKeyName -eq $null }
                foreach ($item in $items) {
                    $out += [PSCustomObject]@{
                        Name       = $item.DisplayName
                        Version    = $item.DisplayVersion
                        Type       = 'Desktop'
                        Identifier = if ($item.QuietUninstallString) { $item.QuietUninstallString } else { $item.UninstallString }
                    }
                }
            } catch { }
        }

        # 2. Modern AppX Packages (Via globally readable AppxAllUserStore registry key)
        # This completely bypasses the strict SYSTEM-only ACLs on the Repository key
        # and sidesteps the WinRM Get-AppxPackage '<Module>' crash entirely.
        try {
            $appxRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications"
            if (Test-Path -LiteralPath $appxRegPath) {
                $appxItems = Get-ChildItem -LiteralPath $appxRegPath -ErrorAction Stop
                
                foreach ($item in $appxItems) {
                    $fullName = $item.PSChildName
                    
                    # Filter out system frameworks, language packs, and neutral UI elements
                    if ($fullName -notmatch "(?i)neutral|language|scale|VCLibs|NET\.Native|DirectX|UI\.Xaml|Services\.Targeting") {
                        $nameParts = $fullName -split '_'
                        $displayName = if ($nameParts.Count -gt 0) { $nameParts[0] } else { $fullName }
                        
                        $out += [PSCustomObject]@{
                            Name       = $displayName
                            Version    = if ($nameParts.Count -gt 1) { $nameParts[1] } else { "" }
                            Type       = 'AppX'
                            Identifier = $fullName
                        }
                    }
                }
            }
        } catch { 
            # Safe catch if the path is completely missing on older OS versions
        }

        return $out | Sort-Object Name -Unique
    }

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop
    } catch {
        Write-Warning "Failed to query software on $($ComputerName): $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Attempts to silently uninstall software remotely (Supports Desktop and AppX).
#>
function Uninstall-HDRemoteSoftware {
    param (
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][string]$Identifier,
        [Parameter(Mandatory=$true)][string]$Type
    )

    $scriptBlock = {
        param($targetId, $targetType)
        
        try {
            if ($targetType -eq 'AppX') {
                # Execute in a discrete, non-interactive powershell process to bypass WinRM WinRT loading exceptions
                # 🛡️ Sentinel: Base64 encode untrusted input to prevent command injection
                $encodedId = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($targetId))
                $psCmd = "`$id = [System.Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedId')); Remove-AppxPackage -Package `$id -AllUsers"
                $encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($psCmd))
                Start-Process powershell.exe -ArgumentList @("-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", $encodedCmd) -Wait -WindowStyle Hidden
            } else {
                $cmd = $targetId
                $exe = ''
                $argsString = ''

                # Safely parse the uninstaller string to separate executable from arguments
                if ($cmd -match '^\s*"([^"]+)"(.*)$') {
                    $exe = $matches[1]
                    $argsString = $matches[2].Trim()
                } else {
                    $possibleExe = ''
                    $remainingArgs = ''
                    $parts = $cmd -split '\s+'

                    for ($i = 0; $i -lt $parts.Count; $i++) {
                        if ($i -gt 0) { $possibleExe += ' ' }
                        $possibleExe += $parts[$i]

                        if ($possibleExe -match '(?i)\.exe$') {
                            $exe = $possibleExe
                            if ($i + 1 -lt $parts.Count) {
                                $remainingArgs = ($parts[($i+1)..($parts.Count-1)]) -join ' '
                            }
                            break
                        }
                    }

                    if (-not $exe) {
                        $parts = $cmd -split '\s+', 2
                        $exe = $parts[0]
                        if ($parts.Count -gt 1) {
                            $remainingArgs = $parts[1]
                        }
                    }
                    $argsString = $remainingArgs
                }

                $argsArray = @()
                if (-not [string]::IsNullOrWhiteSpace($argsString)) {
                    # Split arguments by spaces, keeping quoted strings intact
                    $regex = '(?:[^\s"]|"[^"]*")+'
                    $matchesArg = [regex]::Matches($argsString, $regex)
                    foreach ($m in $matchesArg) {
                        $argsArray += $m.Value
                    }
                }

                # Attempt to convert known installers to silent execution
                if ($exe -match '(?i)msiexec(\.exe)?$') {
                    for ($i = 0; $i -lt $argsArray.Count; $i++) {
                        $argsArray[$i] = $argsArray[$i] -replace '(?i)/I', '/X'
                    }
                    $joinedArgs = $argsArray -join ' '
                    if ($joinedArgs -notmatch '(?i)/q') {
                        $argsArray += '/qn'
                        $argsArray += '/norestart'
                    }
                } elseif ($exe -match '(?i)unins\d{3}\.exe$') {
                    $joinedArgs = $argsArray -join ' '
                    if ($joinedArgs -notmatch '(?i)/VERYSILENT') { $argsArray += '/VERYSILENT' }
                    if ($joinedArgs -notmatch '(?i)/SUPPRESSMSGBOXES') { $argsArray += '/SUPPRESSMSGBOXES' }
                    if ($joinedArgs -notmatch '(?i)/NORESTART') { $argsArray += '/NORESTART' }
                } elseif ($exe -match '(?i)uninstall\.exe$') {
                    $joinedArgs = $argsArray -join ' '
                    if ($joinedArgs -notmatch '(?i)/S') { $argsArray += '/S' }
                }

                # Start the process directly to avoid cmd.exe injection vulnerabilities
                Start-Process -FilePath $exe -ArgumentList $argsArray -Wait -WindowStyle Hidden
            }
            return $true
        } catch {
            throw $_
        }
    }

    try {
        Invoke-Command -ComputerName $ComputerName -ArgumentList $Identifier, $Type -ScriptBlock $scriptBlock -ErrorAction Stop
        return $true
    } catch {
        Write-Warning "Failed to uninstall $Identifier on $($ComputerName): $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Sends a Wake-on-LAN Magic Packet to a specific MAC Address.
#>
function Send-RemoteWakeOnLan {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MacAddress
    )
    try {
        # Clean the MAC format and validate length
        $mac = $MacAddress -replace '[^0-9a-fA-F]', ''
        if ($mac.Length -ne 12) { throw "Invalid MAC Address format. Needs 12 Hex characters." }
        
        # Convert to byte array
        $macBytes = [byte[]]::new(6)
        for ($i = 0; $i -lt 6; $i++) { 
            $macBytes[$i] = [convert]::ToByte($mac.Substring($i * 2, 2), 16) 
        }
        
        # Construct the magic packet (6x 0xFF, followed by 16x MAC Address)
        $packet = [byte[]](@(0xFF) * 6) + ($macBytes * 16)
        
        # Broadcast the UDP packet
        $udp = [System.Net.Sockets.UdpClient]::new()
        $udp.Connect([System.Net.IPAddress]::Broadcast, 9)
        $udp.Send($packet, $packet.Length) | Out-Null
        $udp.Close()
        
        return $true
    } catch {
        Write-Warning "Failed to send Wake-on-LAN packet: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Sets the BIOS/UEFI Administrator or Setup password on supported remote machines.
    Strictly verifies that no password currently exists before executing.
#>
function Set-RemoteBIOSPassword {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][string]$NewPassword
    )
    $scriptBlock = {
        param($pwd)
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        
        # 1 = Enabled. If a password is set, abort operation to prevent locking the motherboard accidentally.
        if ($cs.AdminPasswordStatus -eq 1) {
            throw "A BIOS password is already enabled on this system (Status: 1). This tool is restricted to setting new passwords only."
        }

        $mfg = $cs.Manufacturer
        if ($mfg -match "(?i)LENOVO") {
            try {
                $lenovoSvc = Get-CimInstance -Namespace "root\wmi" -ClassName "Lenovo_SetBiosPassword" -ErrorAction Stop | Select-Object -First 1
                if (-not $lenovoSvc) { throw "Lenovo WMI provider not found." }
                
                $args = @{ CurrentSetting=""; Password=$pwd; PasswordType="sup"; Encoding="ascii" }
                Invoke-CimMethod -InputObject $lenovoSvc -MethodName "SetBiosPassword" -Arguments $args -ErrorAction Stop | Out-Null
                return $true
            } catch { throw "Lenovo WMI Error: $($_.Exception.Message)" }
        }
        elseif ($mfg -match "(?i)Dell") {
            try {
                $dellSvc = Get-CimInstance -Namespace "root\dcim\sysman" -ClassName "DCIM_BIOSService" -ErrorAction Stop | Select-Object -First 1
                if (-not $dellSvc) { throw "Dell Command | Monitor framework is not accessible." }
                
                $args = @{ PasswordType=2; OldPassword=""; NewPassword=$pwd }
                $res = Invoke-CimMethod -InputObject $dellSvc -MethodName "SetPassword" -Arguments $args -ErrorAction Stop
                if ($res.ReturnValue -ne 0) { throw "Dell WMI Method failed with exit code: $($res.ReturnValue)" }
                return $true
            } catch { throw "Dell WMI Error: $($_.Exception.Message)" }
        }
        elseif ($mfg -match "(?i)HP|Hewlett-Packard") {
            try {
                $hpSvc = $null
                # Array of known HP namespaces to seamlessly handle different firmware versions
                $hpNamespaces = @("root\HP\InstrumentedBIOS", "root\HP", "root\HewlettPackard\ComputerSetup")
                
                foreach ($ns in $hpNamespaces) {
                    try {
                        # Query the instance directly to ensure the namespace and class are fully valid
                        $hpSvc = Get-CimInstance -Namespace $ns -ClassName "hp_biossettinginterface" -ErrorAction Stop | Select-Object -First 1
                        if ($hpSvc) { break }
                    } catch { continue }
                }

                if (-not $hpSvc) { 
                    throw "The HP WMI Provider is missing on this computer. Ensure HP Client Management Script Library (CMSL) is installed." 
                }

                # Attempt primary generic Setup Password interface using instance method invocation
                $args = @{ Name="Setup Password"; Value="<utf-16/>$pwd"; Password="" }
                $res = Invoke-CimMethod -InputObject $hpSvc -MethodName "SetBIOSSetting" -Arguments $args -ErrorAction SilentlyContinue
                
                # If the primary interface fails, fallback to standard business model naming
                if (-not $res -or $res.Return -ne 0) {
                    $args2 = @{ Name="Administrator Password"; Value="<utf-16/>$pwd"; Password="" }
                    $res2 = Invoke-CimMethod -InputObject $hpSvc -MethodName "SetBIOSSetting" -Arguments $args2 -ErrorAction Stop
                    if ($res2.Return -ne 0) { throw "Method failed with exit code: $($res2.Return)" }
                }
                return $true
            } catch { throw "HP WMI Error: $($_.Exception.Message)" }
        }
        else {
            throw "The manufacturer '$mfg' is currently not supported for remote BIOS password provisioning."
        }
    }

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $NewPassword -ErrorAction Stop
        return $true
    } catch {
        throw $_.Exception.Message
    }
}

Export-ModuleMember -Function *