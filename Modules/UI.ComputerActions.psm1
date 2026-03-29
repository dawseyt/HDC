# ============================================================================
# UI.ComputerActions.psm1 - Active Directory Computer Interface Logic
# ============================================================================

function Register-ComputerUIEvents {
    param($Window, $Config, $State)

    $AppRoot = Split-Path -Path $PSScriptRoot -Parent

    function Clean-WmiString {
        param([string]$Value)
        if ([string]::IsNullOrEmpty($Value)) { return $Value }
        return [regex]::Replace($Value, '[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '')
    }

    $lvData = $Window.FindName("lvData")
    $tabQuickAsset = $Window.FindName("tabQuickAsset")
    $txtTabSysInfo = $Window.FindName("txtTabSysInfo")
    $lvTabProcesses = $Window.FindName("lvTabProcesses")
    $lvTabServices = $Window.FindName("lvTabServices")
    
    $ctxGPResult = $Window.FindName("ctxGPResult")
    $ctxUptime = $Window.FindName("ctxUptime")
    $ctxViewProcesses = $Window.FindName("ctxViewProcesses")
    $ctxDeviceInfo = $Window.FindName("ctxDeviceInfo")
    $ctxGetLAPS = $Window.FindName("ctxGetLAPS")
    $ctxDetails = $Window.FindName("ctxDetails")
    $ctxPrinterMenu = $Window.FindName("ctxPrinterMenu")
    $ctxPSSession = $Window.FindName("ctxPSSession")
    $ctxPowerMenu = $Window.FindName("ctxPowerMenu")
    $ctxRestartComputer = $Window.FindName("ctxRestartComputer")
    $ctxShutdownComputer = $Window.FindName("ctxShutdownComputer")
    $ctxActiveUsers = $Window.FindName("ctxActiveUsers")
    
    $ctxSoftwareManager = $Window.FindName("MenuItem_SoftwareManager")

    # --- Dynamic Menu Injection (Software Manager) ---
    if ($lvData -and $lvData.ContextMenu) {
        # Check if already generated
        foreach ($item in $lvData.ContextMenu.Items) {
            if ($item -is [System.Windows.Controls.MenuItem]) {
                if ($item.Name -eq "MenuItem_SoftwareManager" -or $item.Header -match "Manage Software") { $ctxSoftwareManager = $item }
            }
        }

        if (-not $ctxSoftwareManager) {
            $ctxSoftwareManager = New-Object System.Windows.Controls.MenuItem
            $ctxSoftwareManager.Name = "MenuItem_SoftwareManager"
            $ctxSoftwareManager.Header = "Manage Software"
            $sIcon = New-Object System.Windows.Controls.TextBlock
            $sIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $sIcon.Text = [char]0xE118
            $sIcon.FontSize = 14
            $sIcon.FontWeight = [System.Windows.FontWeights]::Bold
            $sIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "AccentFill")
            $ctxSoftwareManager.Icon = $sIcon
            [void]($lvData.ContextMenu.Items.Add($ctxSoftwareManager))
        }
    }

    # --- Remote Quick Actions (Dynamic Context Menu) ---
    $ctxQuickActions = $null
    if ($lvData -and $lvData.ContextMenu) {
        $existingQA = $lvData.ContextMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Remote Quick Actions" }
        if (-not $existingQA) {
            $ctxQuickActions = New-Object System.Windows.Controls.MenuItem
            $ctxQuickActions.Header = "Remote Quick Actions"
            
            $qaIcon = New-Object System.Windows.Controls.TextBlock
            $qaIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $qaIcon.Text = [char]0xE765 # Wrench icon
            $qaIcon.FontSize = 14
            $qaIcon.FontWeight = [System.Windows.FontWeights]::Bold
            $qaIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "AccentFill")
            $ctxQuickActions.Icon = $qaIcon
            
            $actions = @(
                [PSCustomObject]@{ Name = "Force Group Policy Update"; IconCode = [char]0xE895; Script = { gpupdate /force } },
                [PSCustomObject]@{ Name = "Restart Print Spooler"; IconCode = [char]0xE749; Script = { Restart-Service -Name Spooler -Force } },
                [PSCustomObject]@{ Name = "Flush DNS Cache"; IconCode = [char]0xE774; Script = { Clear-DnsClientCache -ErrorAction SilentlyContinue; ipconfig /flushdns } },
                [PSCustomObject]@{ Name = "Send Popup Message"; IconCode = [char]0xE8BD; Script = "MSG" },
                [PSCustomObject]@{ Name = "Reboot Computer"; IconCode = [char]0xE777; Script = { Restart-Computer -Force } }
            )

            $dynamicItems = @{}

            foreach ($action in $actions) {
                $mi = New-Object System.Windows.Controls.MenuItem
                $mi.Header = $action.Name
                $mi.Tag = $action.Script
                
                $miIcon = New-Object System.Windows.Controls.TextBlock
                $miIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
                $miIcon.Text = $action.IconCode
                $miIcon.FontSize = 14
                $miIcon.FontWeight = [System.Windows.FontWeights]::Bold
                $miIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "AccentFill")
                $mi.Icon = $miIcon

                $mi.Add_Click({
                    param($sender, $e)
                    if (-not $lvData.SelectedItem -or $lvData.SelectedItem.Type -ne "Computer") { return }
                    
                    $comp = $lvData.SelectedItem.Name
                    $actionName = $sender.Header
                    $scriptToRun = $sender.Tag
                    $colors = Get-FluentThemeColors $State
                    
                    if ($scriptToRun -is [string] -and $scriptToRun -eq "MSG") {
                        $inputXaml = @"
                        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Send Message" Width="400" SizeToContent="Height" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                            <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                                <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                                <StackPanel Margin="20">
                                    <TextBlock Text="Send Message to $comp" FontSize="15" FontWeight="SemiBold" Foreground="{Theme_Fg}" Margin="0,0,0,10"/>
                                    <TextBox x:Name="txtMsg" Height="60" TextWrapping="Wrap" AcceptsReturn="True" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_GridBorder}" Margin="0,0,0,15"/>
                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                        <Button x:Name="btnSend" Content="Send" Width="80" Height="28" Margin="0,0,8,0" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsDefault="True"/>
                                        <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" IsCancel="True"/>
                                    </StackPanel>
                                </StackPanel>
                            </Border>
                        </Window>
"@
                        $iXaml = $inputXaml; foreach ($key in $colors.Keys) { $iXaml = $iXaml.Replace("{Theme_$key}", $colors[$key]) }
                        $iReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($iXaml))
                        $inWin = [System.Windows.Markup.XamlReader]::Load($iReader); $inWin.Owner = $Window
                        
                        $inWin.Add_Loaded({ $inWin.FindName("txtMsg").Focus() | Out-Null }.GetNewClosure())

                        $inWin.FindName("btnSend").Add_Click({ $inWin.DialogResult = $true }.GetNewClosure())
                        $inWin.FindName("btnCancel").Add_Click({ $inWin.DialogResult = $false }.GetNewClosure())
                        
                        if ($inWin.ShowDialog() -ne $true) { return }
                        
                        $msgText = $inWin.FindName("txtMsg").Text
                        if ([string]::IsNullOrWhiteSpace($msgText)) { return }
                        
                        $realScript = { param($m) msg.exe * /time:0 $m }
                        $job = Invoke-Command -ComputerName $comp -ScriptBlock $realScript -ArgumentList $msgText -AsJob -ErrorAction SilentlyContinue
                    } else {
                        $confirm = Show-AppMessageBox -Message "Are you sure you want to run '$actionName' on computer '$comp'?" -Title "Confirm Action" -ButtonType "YesNo" -IconType "Question" -OwnerWindow $Window -ThemeColors $colors
                        if ($confirm -ne "Yes") { return }
                        
                        $job = Invoke-Command -ComputerName $comp -ScriptBlock $scriptToRun -AsJob -ErrorAction SilentlyContinue
                    }

                    $loadingXaml = @"
                    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="320" Height="150" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                        <Window.Resources><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Window.Resources>
                        <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                            <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                            <Grid>
                                <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                <StackPanel Grid.Row="0" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0,10,0,15">
                                    <TextBlock Text="Executing $actionName on $comp..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                                    <ProgressBar IsIndeterminate="True" Width="240" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                                </StackPanel>
                                <Button Grid.Row="1" x:Name="btnCancelJob" Content="Cancel" Width="80" Height="26" Margin="0,0,0,10" HorizontalAlignment="Center" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1"/>
                            </Grid>
                        </Border>
                    </Window>
"@
                    $xamlText = $loadingXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                    $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 

                    $timer = New-Object System.Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                    $timerTick = {
                        if ($job.State -ne 'Running') {
                            $timer.Stop(); $loadWin.Close()
                            if ($job.State -eq 'Completed') {
                                Receive-Job $job -ErrorAction SilentlyContinue | Out-Null
                                Show-AppMessageBox -Message "Action '$actionName' completed successfully on $comp." -Title "Success" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                                Add-AppLog -Event "Quick Action" -Username $comp -Details "Executed: $actionName" -Config $Config -State $State -Status "Success" -Color "Green"
                            } else {
                                $reason = if ($job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Host unreachable, Access Denied, or job canceled." }
                                Show-AppMessageBox -Message "Action '$actionName' failed on ${comp}:`n$reason" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                                Add-AppLog -Event "Quick Action" -Username $comp -Details "Failed: $actionName" -Config $Config -State $State -Status "Error" -Color "Red"
                            }
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                        }
                    }.GetNewClosure()
                    $timer.Add_Tick($timerTick)

                    $btnCancelJob = $loadWin.FindName("btnCancelJob")
                    if ($btnCancelJob) {
                        $btnCancelJob.Add_Click({
                            $timer.Stop()
                            try {
                                if ($job -and $job.State -eq 'Running') { Stop-Job $job -Force -ErrorAction SilentlyContinue }
                                Remove-Job $job -Force -ErrorAction SilentlyContinue
                            } catch {}
                            $loadWin.Close()
                        }.GetNewClosure())
                    }

                    $timer.Start()
                }.GetNewClosure())
                
                $dynamicItems[$action.Name] = $mi
            }

            [void]($ctxQuickActions.Items.Add($dynamicItems["Force Group Policy Update"]))
            [void]($ctxQuickActions.Items.Add($dynamicItems["Restart Print Spooler"]))
            [void]($ctxQuickActions.Items.Add($dynamicItems["Flush DNS Cache"]))
            [void]($ctxQuickActions.Items.Add($dynamicItems["Send Popup Message"]))
            
            [void]($ctxQuickActions.Items.Add([System.Windows.Controls.Separator]::new()))
            
            $itemsToMove = @(
                @{ Item = $ctxPSSession; Icon = [char]0xE756 },
                @{ Item = $ctxGPResult; Icon = [char]0xE8A5 },
                @{ Item = $ctxUptime; Icon = [char]0xE91E }
            )
            foreach ($obj in $itemsToMove) {
                $item = $obj.Item
                if ($null -ne $item) {
                    $iconTb = New-Object System.Windows.Controls.TextBlock
                    $iconTb.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
                    $iconTb.Text = $obj.Icon
                    $iconTb.FontSize = 14
                    $iconTb.FontWeight = [System.Windows.FontWeights]::Bold
                    $iconTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "AccentFill")
                    $item.Icon = $iconTb

                    $parent = $item.Parent
                    if ($parent -is [System.Windows.Controls.ItemsControl]) {
                        $parent.Items.Remove($item)
                    } else {
                        if ($lvData.ContextMenu.Items.Contains($item)) { $lvData.ContextMenu.Items.Remove($item) }
                        if ($ctxPowerMenu -and $ctxPowerMenu.Items.Contains($item)) { $ctxPowerMenu.Items.Remove($item) }
                    }
                    [void]($ctxQuickActions.Items.Add($item))
                }
            }

            [void]($ctxQuickActions.Items.Add([System.Windows.Controls.Separator]::new()))
            [void]($ctxQuickActions.Items.Add($dynamicItems["Reboot Computer"]))

            if ($ctxShutdownComputer) {
                $sIcon = New-Object System.Windows.Controls.TextBlock
                $sIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
                $sIcon.Text = [char]0xE7E8
                $sIcon.FontSize = 14
                $sIcon.FontWeight = [System.Windows.FontWeights]::Bold
                $sIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "AccentFill")
                $ctxShutdownComputer.Icon = $sIcon
            }

            $powerItemsToMove = @($ctxShutdownComputer)
            foreach ($item in $powerItemsToMove) {
                if ($null -ne $item) {
                    $parent = $item.Parent
                    if ($parent -is [System.Windows.Controls.ItemsControl]) {
                        $parent.Items.Remove($item)
                    } else {
                        if ($lvData.ContextMenu.Items.Contains($item)) { $lvData.ContextMenu.Items.Remove($item) }
                        if ($ctxPowerMenu -and $ctxPowerMenu.Items.Contains($item)) { $ctxPowerMenu.Items.Remove($item) }
                    }
                    [void]($ctxQuickActions.Items.Add($item))
                }
            }

            if ($ctxPowerMenu) {
                $parent = $ctxPowerMenu.Parent
                if ($parent -is [System.Windows.Controls.ItemsControl]) {
                    $parent.Items.Remove($ctxPowerMenu)
                } elseif ($lvData.ContextMenu.Items.Contains($ctxPowerMenu)) {
                    $lvData.ContextMenu.Items.Remove($ctxPowerMenu)
                }
            }
            if ($ctxRestartComputer) {
                $parent = $ctxRestartComputer.Parent
                if ($parent -is [System.Windows.Controls.ItemsControl]) { $parent.Items.Remove($ctxRestartComputer) }
            }

            [void]($lvData.ContextMenu.Items.Add([System.Windows.Controls.Separator]::new()))
            [void]($lvData.ContextMenu.Items.Add($ctxQuickActions))
        } else {
            $ctxQuickActions = $existingQA
        }
    }

    if ($lvData -and $lvData.ContextMenu) {
        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            $isComp = ($sel -and $sel.Type -eq "Computer")
            $vis = if ($isComp) { "Visible" } else { "Collapsed" }
            
            if ($ctxGPResult)        { $ctxGPResult.Visibility        = $vis }
            if ($ctxUptime)          { $ctxUptime.Visibility          = $vis }
            if ($ctxViewProcesses)   { $ctxViewProcesses.Visibility    = $vis }
            if ($ctxDeviceInfo)      { $ctxDeviceInfo.Visibility      = $vis }
            if ($ctxGetLAPS)         { $ctxGetLAPS.Visibility         = $vis }
            if ($ctxPrinterMenu)     { $ctxPrinterMenu.Visibility     = $vis }
            if ($ctxPSSession)       { $ctxPSSession.Visibility       = $vis }
            if ($ctxPowerMenu)       { $ctxPowerMenu.Visibility       = $vis }
            if ($ctxActiveUsers)     { $ctxActiveUsers.Visibility     = $vis }
            if ($ctxQuickActions)    { $ctxQuickActions.Visibility    = $vis }
            if ($ctxSoftwareManager) { $ctxSoftwareManager.Visibility = $vis }
            
            if ($ctxDetails -and $sel) { $ctxDetails.Visibility = "Visible" }
            elseif ($ctxDetails)       { $ctxDetails.Visibility = "Collapsed" }
            
            if ($isComp -and $ctxActiveUsers) {
                $ctxActiveUsers.Tag = $null
                $ctxActiveUsers.Items.Clear()
                $dummyItem = New-Object System.Windows.Controls.MenuItem
                $dummyItem.Header = "Loading..."
                [void]($ctxActiveUsers.Items.Add($dummyItem))
            }
        }.GetNewClosure())
    }

    # --- GPResult ---
    if ($ctxGPResult) {
        $ctxGPResult.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $comp = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $loadingXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="340" Height="150" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Window.Resources><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Window.Resources>
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                        <Grid>
                            <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0,10,0,15">
                                <TextBlock Text="Querying GPResult XML from $comp..." FontSize="13" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                                <ProgressBar IsIndeterminate="True" Width="260" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                            </StackPanel>
                            <Button Grid.Row="1" x:Name="btnCancelJob" Content="Cancel" Width="80" Height="26" Margin="0,0,0,10" HorizontalAlignment="Center" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1"/>
                        </Grid>
                    </Border>
                </Window>
"@
                $xamlText = $loadingXaml; foreach ($k in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$k}", $colors[$k]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 
                
                $job = Start-Job -ScriptBlock {
                    param($c)
                    try {
                        # Run GPResult locally on the target machine via WinRM
                        $rawXml = Invoke-Command -ComputerName $c -ScriptBlock {
                            $remoteTemp = Join-Path $env:TEMP "gpresult_$([guid]::NewGuid()).xml"
                            
                            # Check if a user is currently logged on to retrieve User-scoped policies
                            $loggedOnUser = $null
                            try { $loggedOnUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName } catch {}
                            
                            if ([string]::IsNullOrWhiteSpace($loggedOnUser)) {
                                & gpresult.exe /SCOPE COMPUTER /X $remoteTemp /F *>&1 | Out-Null
                            } else {
                                & gpresult.exe /USER $loggedOnUser /X $remoteTemp /F *>&1 | Out-Null
                            }
                            
                            if (Test-Path -LiteralPath $remoteTemp) {
                                $xmlContent = [System.IO.File]::ReadAllText($remoteTemp, [System.Text.Encoding]::Unicode)
                                Remove-Item -LiteralPath $remoteTemp -Force -ErrorAction SilentlyContinue
                                return $xmlContent
                            }
                            throw "Failed to generate GPResult XML on target computer."
                        } -ErrorAction Stop
                        
                        if ([string]::IsNullOrWhiteSpace($rawXml)) { throw "Returned XML is empty." }

                        # Return the flat XML string to perfectly bypass background job serialization depths
                        return [PSCustomObject]@{ Success = $true; Xml = $rawXml }
                    } catch {
                        return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message }
                    }
                } -ArgumentList $comp
                
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                $startTime = Get-Date

                $timerTick = {
                    # Add a robust 60-second timeout
                    if ($job.State -ne 'Running' -or ((Get-Date) - $startTime).TotalSeconds -ge 60) {
                        $timer.Stop(); $loadWin.Close()
                        
                        if ($job.State -eq 'Completed') {
                            $res = Receive-Job $job -ErrorAction SilentlyContinue
                            Remove-Job $job -Force -ErrorAction SilentlyContinue

                            if ($res -and $res.Success) { 
                                try {
                                    # PARSE XML ON THE MAIN UI THREAD
                                    [xml]$xml = $res.Xml
                                    
                                    function Parse-GPONodes {
                                        param($nodes, $scope, $xmlDoc)
                                        $out = @()
                                        if (-not $nodes) { return $out }
                                        
                                        foreach ($node in $nodes) {
                                            $nameNode = $node.SelectSingleNode("*[local-name()='Name']")
                                            $gpoName = if ($nameNode) { $nameNode.InnerText } else { "Unknown" }

                                            $filterNode = $node.SelectSingleNode("*[local-name()='FilterAllowed']")
                                            $isApplied = ($filterNode -and $filterNode.InnerText.Trim().ToLower() -eq 'true')

                                            $status = "Applied"
                                            if (-not $isApplied) {
                                                $reasonNode = $node.SelectSingleNode("*[local-name()='FilterDeniedReason']")
                                                $reason = if ($reasonNode) { $reasonNode.InnerText } else { "Unknown" }
                                                $status = "Denied: $reason"
                                            }

                                            $orderNode = $node.SelectSingleNode("*[local-name()='Link']/*[local-name()='AppliedOrder']")
                                            $link = if ($orderNode) { $orderNode.InnerText } else { "-" }

                                            $idNode = $node.SelectSingleNode("*[local-name()='Identifier']")
                                            $pathNode = $node.SelectSingleNode("*[local-name()='Path']")
                                            $adPath = if ($pathNode) { $pathNode.InnerText.Trim() } else { "Unknown" }
                                            
                                            # Regex scrape the GUID
                                            $guid = "Unknown"
                                            if ($idNode -and $idNode.InnerText -match '\{[a-fA-F0-9-]{36}\}') { $guid = $matches[0] }
                                            elseif ($adPath -match '\{[a-fA-F0-9-]{36}\}') { $guid = $matches[0] }
                                            elseif ($gpoName -match '\{[a-fA-F0-9-]{36}\}') { $guid = $matches[0] }

                                            # Only return basic fields for the DataGrid.
                                            # Advanced properties (Drives, Printers, etc.) will be dynamically 
                                            # mapped into a TreeView on double-click.
                                            $out += [PSCustomObject]@{
                                                Scope     = $scope
                                                Name      = $gpoName
                                                Status    = $status
                                                LinkOrder = $link
                                                IsApplied = $isApplied
                                                Id        = $guid
                                                Path      = $adPath
                                            }
                                        }
                                        return $out
                                    }

                                    $gpoList = @()
                                    $gpoList += Parse-GPONodes -nodes ($xml.SelectNodes("//*[local-name()='ComputerResults']/*[local-name()='GPO']")) -scope "Computer" -xmlDoc $xml
                                    $gpoList += Parse-GPONodes -nodes ($xml.SelectNodes("//*[local-name()='UserResults']/*[local-name()='GPO']")) -scope "User" -xmlDoc $xml
                                    
                                    # Sort safely: Protect against casting "-" to [int]
                                    $gpoList = $gpoList | Sort-Object @{Expression={$_.IsApplied}; Descending=$true}, @{Expression={$_.Scope}; Ascending=$true}, @{Expression={if ($_.LinkOrder -match '\d') { [int]$_.LinkOrder } else { 9999 }}; Ascending=$true}
                                    
                                    # Build a Native WPF Window to display the parsed GPOs
                                    $gpoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="GPO Results" Width="700" Height="500" MinWidth="500" MinHeight="350" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
    <Window.Resources>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="RowBackground" Value="Transparent"/>
            <Setter Property="Foreground" Value="{Theme_Fg}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{Theme_BtnBg}"/><Setter Property="Foreground" Value="{Theme_SecFg}"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Padding" Value="10,6"/><Setter Property="BorderThickness" Value="0,0,0,1"/><Setter Property="BorderBrush" Value="{Theme_GridBorder}"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="Transparent"/><Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="{Theme_PrimaryBg}"/><Setter Property="Foreground" Value="{Theme_PrimaryFg}"/></Trigger></Style.Triggers>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
        </Style>
    </Window.Resources>
    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
        <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3" Color="Black"/></Border.Effect>
        <Grid>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="20,18,20,12" Cursor="Hand" BorderThickness="0,0,0,1" BorderBrush="{Theme_BtnBorder}">
                <Grid>
                    <StackPanel>
                        <TextBlock Text="Group Policy Results" FontSize="16" FontWeight="SemiBold" Foreground="{Theme_Fg}"/>
                        <TextBlock Text="Target: $comp" FontSize="12" Foreground="{Theme_SecFg}" Margin="0,3,0,0"/>
                    </StackPanel>
                    <TextBlock Text="&#x2715;" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="{Theme_SecFg}" FontSize="16" Cursor="Hand" x:Name="btnXClose" ToolTip="Close"/>
                </Grid>
            </Border>
            
            <DataGrid x:Name="dgGPOs" Grid.Row="1" Margin="10" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" ToolTip="Double-click a policy to view extended details.">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Scope" Binding="{Binding Scope}" Width="70"/>
                    <DataGridTextColumn Header="Order" Binding="{Binding LinkOrder}" Width="60"/>
                    <DataGridTextColumn Header="GPO Name" Binding="{Binding Name}" Width="*"/>
                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="200"/>
                </DataGrid.Columns>
            </DataGrid>

            <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}">
                <Grid>
                    <Button x:Name="btnClose" Content="Close" Width="80" Height="28" HorizontalAlignment="Right" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsCancel="True"/>
                    <Thumb x:Name="resizeGrip" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="12" Height="12" Cursor="SizeNWSE" Background="Transparent" Margin="0,0,-2,-2"/>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
"@
                                    $gx = $gpoXaml; foreach ($k in $colors.Keys) { $gx = $gx.Replace("{Theme_$k}", $colors[$k]) }
                                    $gpoWin = [System.Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new($gx)))
                                    $gpoWin.Owner = $Window
                                    $gpoWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $gpoWin.DragMove() }.GetNewClosure())
                                    $gpoWin.FindName("btnClose").Add_Click({ $gpoWin.Close() }.GetNewClosure())
                                    
                                    $xClose = $gpoWin.FindName("btnXClose")
                                    if ($xClose) { $xClose.Add_PreviewMouseLeftButtonDown({ param($s,$e) $e.Handled = $true; $gpoWin.Close() }.GetNewClosure()) }
                                    
                                    $resizeGrip = $gpoWin.FindName("resizeGrip")
                                    if ($resizeGrip) {
                                        $resizeGrip.Add_DragDelta({ param($s,$e)
                                            $nw = [math]::Max(500, $gpoWin.Width + $e.HorizontalChange)
                                            $nh = [math]::Max(350, $gpoWin.Height + $e.VerticalChange)
                                            $gpoWin.Width = $nw; $gpoWin.Height = $nh
                                        }.GetNewClosure())
                                    }

                                    $dg = $gpoWin.FindName("dgGPOs")
                                    if ($dg -and $gpoList) { 
                                        $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[psobject]]::new($gpoList) 
                                        
                                        # Double-click Recursive XML Drill Down Event
                                        $dg.Add_MouseDoubleClick({
                                            param($s, $e)
                                            if ($dg.SelectedItem) {
                                                $selGpo = $dg.SelectedItem
                                                
                                                # Safely re-fetch theme colors within this nested closure
                                                $drillColors = Get-FluentThemeColors $State
                                                
                                                $detXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="GPO Details" Width="600" Height="600" MinWidth="450" MinHeight="400" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="TreeViewItem">
            <Setter Property="IsExpanded" Value="True"/>
            <Setter Property="Margin" Value="0,2,0,0"/>
        </Style>
    </Window.Resources>
    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
        <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3" Color="Black"/></Border.Effect>
        <Grid>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <Border x:Name="TitleBarDet" Grid.Row="0" Background="Transparent" Padding="20,18,20,12" BorderThickness="0,0,0,1" BorderBrush="{Theme_BtnBorder}" Cursor="Hand">
                <Grid>
                    <StackPanel>
                        <TextBlock Text="Advanced Policy Details" FontSize="16" FontWeight="SemiBold" Foreground="{Theme_Fg}"/>
                        <TextBlock x:Name="lblDetSub" Text="Loading..." FontSize="12" Foreground="{Theme_SecFg}" Margin="0,3,0,0"/>
                    </StackPanel>
                </Grid>
            </Border>
            
            <Grid Grid.Row="1" Margin="20,15,20,15">
                <Border Background="{Theme_BtnBg}" BorderBrush="{Theme_GridBorder}" BorderThickness="1" CornerRadius="4">
                    <TreeView x:Name="tvDetails" Background="Transparent" BorderThickness="0" Padding="8"/>
                </Border>
            </Grid>

            <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}">
                <Grid>
                    <Button x:Name="btnCloseDet" Content="Close" Width="80" Height="28" HorizontalAlignment="Right" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsCancel="True"/>
                    <Thumb x:Name="resizeGripDet" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="12" Height="12" Cursor="SizeNWSE" Background="Transparent" Margin="0,0,-2,-2"/>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
"@
                                                $dx = $detXaml; foreach ($k in $drillColors.Keys) { $dx = $dx.Replace("{Theme_$k}", $drillColors[$k]) }
                                                $detWin = [System.Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dx)))
                                                $detWin.Owner = $gpoWin
                                                
                                                $titleBarDet = $detWin.FindName("TitleBarDet")
                                                if ($titleBarDet) { $titleBarDet.Add_MouseLeftButtonDown({ $detWin.DragMove() }.GetNewClosure()) }
                                                
                                                $resizeGripDet = $detWin.FindName("resizeGripDet")
                                                if ($resizeGripDet) {
                                                    $resizeGripDet.Add_DragDelta({ param($s,$ev)
                                                        $nw = [math]::Max(450, $detWin.Width + $ev.HorizontalChange)
                                                        $nh = [math]::Max(400, $detWin.Height + $ev.VerticalChange)
                                                        $detWin.Width = $nw; $detWin.Height = $nh
                                                    }.GetNewClosure())
                                                }
                                                
                                                $detWin.FindName("lblDetSub").Text = "$($selGpo.Name)  |  $($selGpo.Scope) Scope"
                                                
                                                $tvDetails = $detWin.FindName("tvDetails")
                                                if ($tvDetails) {
                                                    # Recursive function to build TreeViewItems from XML Elements
                                                    function Build-Tree {
                                                        param($node)
                                                        if ($node -is [System.Xml.XmlDeclaration]) { return $null }
                                                        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Text) { return $null }
                                                        
                                                        $tvi = New-Object System.Windows.Controls.TreeViewItem
                                                        $headerText = $node.LocalName
                                                        
                                                        if ($node.HasAttributes) {
                                                            $attrs = @()
                                                            foreach ($a in $node.Attributes) { $attrs += "$($a.Name)='$($a.Value)'" }
                                                            $headerText += " [" + ($attrs -join ', ') + "]"
                                                        }
                                                        
                                                        if ($node.HasChildNodes -and $node.ChildNodes.Count -eq 1 -and $node.FirstChild.NodeType -eq [System.Xml.XmlNodeType]::Text) {
                                                            $headerText += " : " + $node.FirstChild.InnerText.Trim()
                                                        }
                                                        
                                                        $tb = New-Object System.Windows.Controls.TextBlock
                                                        $tb.Text = $headerText
                                                        $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($drillColors.Fg))
                                                        $tvi.Header = $tb
                                                        
                                                        if ($node.HasChildNodes) {
                                                            foreach ($c in $node.ChildNodes) {
                                                                $childTvi = Build-Tree -node $c
                                                                if ($childTvi) { [void]($tvi.Items.Add($childTvi)) }
                                                            }
                                                        }
                                                        return $tvi
                                                    }
                                                    
                                                    $guid = $selGpo.Id
                                                    if ($guid -eq "Unknown") {
                                                        $emptyTvi = New-Object System.Windows.Controls.TreeViewItem
                                                        $emptyTb = New-Object System.Windows.Controls.TextBlock
                                                        $emptyTb.Text = "Cannot map XML tree. No valid GUID identifier found for this policy."
                                                        $emptyTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($drillColors.Fg))
                                                        $emptyTvi.Header = $emptyTb
                                                        [void]($tvDetails.Items.Add($emptyTvi))
                                                    } else {
                                                        # Query main GPO Node
                                                        $gpoNode = $xml.SelectSingleNode("//*[local-name()='GPO'][*[local-name()='Identifier'][text()='$guid']]")
                                                        if ($gpoNode) {
                                                            $gpoTvi = Build-Tree -node $gpoNode
                                                            $gpoTvi.IsExpanded = $true
                                                            [void]($tvDetails.Items.Add($gpoTvi))
                                                        }
                                                        
                                                        # Query ExtensionData Nodes tied to this GPO
                                                        $extNodes = $xml.SelectNodes("//*[local-name()='ExtensionData'][.//*[local-name()='Identifier'][text()='$guid']]")
                                                        if ($extNodes) {
                                                            foreach ($ext in $extNodes) {
                                                                $extTvi = Build-Tree -node $ext
                                                                $extTvi.IsExpanded = $true
                                                                [void]($tvDetails.Items.Add($extTvi))
                                                            }
                                                        }
                                                        
                                                        if ($tvDetails.Items.Count -eq 0) {
                                                            $emptyTvi = New-Object System.Windows.Controls.TreeViewItem
                                                            $emptyTb = New-Object System.Windows.Controls.TextBlock
                                                            $emptyTb.Text = "No advanced XML properties found for this GPO."
                                                            $emptyTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($drillColors.Fg))
                                                            $emptyTvi.Header = $emptyTb
                                                            [void]($tvDetails.Items.Add($emptyTvi))
                                                        }
                                                    }
                                                }
                                                
                                                $detWin.FindName("btnCloseDet").Add_Click({ $detWin.Close() }.GetNewClosure())
                                                $detWin.ShowDialog() | Out-Null
                                            }
                                        }.GetNewClosure())
                                    }
                                    
                                    Show-CenteredOnOwner -ChildWindow $gpoWin -OwnerWindow $Window
                                    $gpoWin.Show()
                                } catch {
                                    Show-AppMessageBox -Message "Failed to render GPResult:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                                }
                            } 
                            else {
                                $reason = if ($res -and $res.Error) { $res.Error } elseif ($job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Host unreachable, Access Denied, or job canceled." }
                                Show-AppMessageBox -Message "GPResult failed:`n$reason" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                            }
                        } else {
                            Stop-Job $job -ErrorAction SilentlyContinue
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                            $reason = if ($job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Job timed out after 60 seconds." }
                            Show-AppMessageBox -Message "GPResult failed:`n$reason" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                        }
                    }
                }.GetNewClosure()
                $timer.Add_Tick($timerTick)

                $btnCancelJob = $loadWin.FindName("btnCancelJob")
                if ($btnCancelJob) {
                    $btnCancelJob.Add_Click({
                        $timer.Stop()
                        try {
                            if ($job -and $job.State -eq 'Running') { Stop-Job $job -Force -ErrorAction SilentlyContinue }
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                        } catch {}
                        $loadWin.Close()
                    }.GetNewClosure())
                }

                $timer.Start()
            }
        }.GetNewClosure())
    }

    # --- Uptime ---
    if ($ctxUptime) {
        $ctxUptime.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $computer = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $loadingXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="320" Height="150" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Window.Resources><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Window.Resources>
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                        <Grid>
                            <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0,10,0,15">
                                <TextBlock Text="Querying Uptime for $computer..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                                <ProgressBar IsIndeterminate="True" Width="240" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                            </StackPanel>
                            <Button Grid.Row="1" x:Name="btnCancelJob" Content="Cancel" Width="80" Height="26" Margin="0,0,0,10" HorizontalAlignment="Center" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1"/>
                        </Grid>
                    </Border>
                </Window>
"@
                $xamlText = $loadingXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 

                $job = Start-Job -ScriptBlock {
                    param($c)
                    try {
                        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $c -ErrorAction Stop
                        $lastBoot = $os.LastBootUpTime; $uptime = (Get-Date) - $lastBoot
                        return [PSCustomObject]@{ Success = $true; LastBootUpTime = $lastBoot; Days = $uptime.Days; Hours = $uptime.Hours; Minutes = $uptime.Minutes }
                    } catch { return [PSCustomObject]@{ Success = $false; ErrorMessage = $_.Exception.Message } }
                } -ArgumentList $computer

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500); $startTime = Get-Date
                
                $timerTick = {
                    if ($job.State -ne 'Running' -or ((Get-Date) - $startTime).TotalSeconds -ge 10) {
                        $timer.Stop(); $loadWin.Close()
                        if ($job.State -eq 'Completed') {
                            $up = Receive-Job $job -ErrorAction SilentlyContinue
                            if ($up -and $up.Success) { Show-AppMessageBox -Message "Computer: $computer`n`nStatus: ONLINE`nLast Boot: $($up.LastBootUpTime)`n`nUptime: $($up.Days) days, $($up.Hours) hours, $($up.Minutes) minutes" -Title "Uptime Check" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors | Out-Null } 
                            elseif ($up -and -not $up.Success) { Show-AppMessageBox -Message "Failed to query uptime for $computer.`n`nError: $($up.ErrorMessage)" -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null } 
                            else { Show-AppMessageBox -Message "Failed to retrieve uptime data." -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null }
                        } else {
                            Stop-Job $job -ErrorAction SilentlyContinue
                            $reason = "Connection timed out after 10 seconds or the job crashed."
                            if ($job.ChildJobs[0].JobStateInfo.Reason) { $reason = $job.ChildJobs[0].JobStateInfo.Reason.Message }
                            Show-AppMessageBox -Message "Failed to query uptime for $computer.`n`nError: $reason" -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                        }
                        Remove-Job $job -Force -ErrorAction SilentlyContinue
                    }
                }.GetNewClosure()
                $timer.Add_Tick($timerTick)

                $btnCancelJob = $loadWin.FindName("btnCancelJob")
                if ($btnCancelJob) {
                    $btnCancelJob.Add_Click({
                        $timer.Stop()
                        try {
                            if ($job -and $job.State -eq 'Running') { Stop-Job $job -Force -ErrorAction SilentlyContinue }
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                        } catch {}
                        $loadWin.Close()
                    }.GetNewClosure())
                }

                $timer.Start()
            }
        }.GetNewClosure())
    }

    # --- Device Info (Quick Popup from main list right-click) ---
    if ($ctxDeviceInfo) {
        $ctxDeviceInfo.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $computer = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State

                $loadingXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Device Info" Width="480" Height="160" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Window.Resources><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Window.Resources>
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                        <Grid>
                            <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0,10,0,15">
                                <TextBlock Text="Querying hardware info for $computer..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                                <ProgressBar IsIndeterminate="True" Width="240" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                            </StackPanel>
                            <Button Grid.Row="1" x:Name="btnCancelJob" Content="Cancel" Width="80" Height="26" Margin="0,0,0,10" HorizontalAlignment="Center" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1"/>
                        </Grid>
                    </Border>
                </Window>
"@
                $xamlText = $loadingXaml
                foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $loadWin = [System.Windows.Markup.XamlReader]::Load($reader)
                $loadWin.Owner = $Window
                $loadWin.Show()

                Add-AppLog -Event "Query" -Username "System" -Details "Querying device info for $computer..." -Config $Config -State $State -Status "Info"

                $job = Start-Job -ScriptBlock {
                    param($c)
                    try {
                        $result = Invoke-Command -ComputerName $c -ScriptBlock {
                            function CleanStr($v) {
                                if ([string]::IsNullOrEmpty($v)) { return $v }
                                return [regex]::Replace($v, '[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '')
                            }
                            $cs  = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
                            $bio = Get-CimInstance Win32_BIOS            -ErrorAction Stop
                            $bat = Get-CimInstance Win32_Battery          -ErrorAction SilentlyContinue
                            $model = if ($cs.Manufacturer -match "LENOVO" -and $cs.Model.Length -ge 4) { $cs.Model.Substring(0,4) } else { $cs.Model }
                            $batteryStatus = if ($bat) { "$($bat.EstimatedChargeRemaining)%" } else { "No Battery / Desktop" }
                            
                            $adminPwdStatus = switch ($cs.AdminPasswordStatus) {
                                0 { "Disabled" }
                                1 { "Enabled" }
                                2 { "Not Implemented" }
                                3 { "Unknown" }
                                default { if ($null -ne $cs.AdminPasswordStatus) { "Unknown ($($cs.AdminPasswordStatus))" } else { "Unknown" } }
                            }

                            return [PSCustomObject]@{
                                ComputerName        = CleanStr $cs.Name
                                Manufacturer        = CleanStr $cs.Manufacturer
                                Model               = CleanStr $model
                                SystemFamily        = CleanStr $cs.SystemFamily
                                SerialNumber        = CleanStr $bio.SerialNumber
                                BIOSVersion         = CleanStr $bio.SMBIOSBIOSVersion
                                BatteryStatus       = CleanStr $batteryStatus
                                AdminPasswordStatus = $adminPwdStatus
                                QueryTime           = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                            }
                        } -ErrorAction Stop
                        return [PSCustomObject]@{ Success = $true
                            ComputerName        = $result.ComputerName
                            Manufacturer        = $result.Manufacturer
                            Model               = $result.Model
                            SystemFamily        = $result.SystemFamily
                            SerialNumber        = $result.SerialNumber
                            BIOSVersion         = $result.BIOSVersion
                            BatteryStatus       = $result.BatteryStatus
                            AdminPasswordStatus = $result.AdminPasswordStatus
                            QueryTime           = $result.QueryTime
                        }
                    } catch {
                        return [PSCustomObject]@{ Success = $false; ErrorMessage = $_.Exception.Message }
                    }
                } -ArgumentList $computer

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                $startTime = Get-Date

                $timerTick = {
                    if ($job.State -ne 'Running' -or ((Get-Date) - $startTime).TotalSeconds -ge 15) {
                        $timer.Stop(); $loadWin.Close()
                        if ($job.State -eq 'Completed') {
                            $info = Receive-Job $job -ErrorAction SilentlyContinue
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                            if ($info -and $info.Success) {
                                $diXaml = @"
                                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                                        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                                        Title="Device Info" Width="480" SizeToContent="Height"
                                        WindowStartupLocation="CenterOwner" WindowStyle="None"
                                        AllowsTransparency="True" Background="Transparent"
                                        FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                                    <Window.Resources>
                                        <Style TargetType="Button">
                                            <Setter Property="Template">
                                                <Setter.Value>
                                                    <ControlTemplate TargetType="Button">
                                                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                        </Border>
                                                        <ControlTemplate.Triggers>
                                                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                                                        </ControlTemplate.Triggers>
                                                    </ControlTemplate>
                                                </Setter.Value>
                                            </Setter>
                                        </Style>
                                    </Window.Resources>
                                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                                        <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3" Color="Black"/></Border.Effect>
                                        <Grid>
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="20,18,20,0" Cursor="Hand">
                                                <StackPanel>
                                                    <TextBlock Text="Device Info" FontSize="16" FontWeight="SemiBold" Foreground="{Theme_Fg}"/>
                                                    <TextBlock Text="$computer" FontSize="12" Foreground="{Theme_SecFg}" Margin="0,3,0,0"/>
                                                </StackPanel>
                                            </Border>
                                            <StackPanel Grid.Row="1" Margin="20,14,20,4">
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_BtnBg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Computer Name" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.ComputerName)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_Bg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Manufacturer" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.Manufacturer)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_BtnBg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Model" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.Model)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_Bg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="System Family" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.SystemFamily)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_BtnBg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Serial Number" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.SerialNumber)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_Bg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="BIOS Version" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.BIOSVersion)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_BtnBg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Battery Status" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.BatteryStatus)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,5" Padding="14,9" Background="{Theme_Bg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Admin Pwd Status" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="13" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.AdminPasswordStatus)" Foreground="{Theme_Fg}" FontSize="13" VerticalAlignment="Center"/></Grid>
                                                </Border>
                                                <Border Margin="0,0,0,0" Padding="14,9" Background="{Theme_BtnBg}" CornerRadius="6" BorderBrush="{Theme_GridBorder}" BorderThickness="1">
                                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <TextBlock Grid.Column="0" Text="Queried At" FontWeight="SemiBold" Foreground="{Theme_SecFg}" FontSize="12" FontStyle="Italic" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="$($info.QueryTime)" Foreground="{Theme_SecFg}" FontSize="12" FontStyle="Italic" VerticalAlignment="Center"/></Grid>
                                                </Border>
                                            </StackPanel>
                                            <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}" Margin="0,16,0,0">
                                                <Button x:Name="btnDiOk" Content="OK" Width="80" Height="28" HorizontalAlignment="Right" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsDefault="True"/>
                                            </Border>
                                        </Grid>
                                    </Border>
                                </Window>
"@
                                $diText = $diXaml
                                foreach ($key in $colors.Keys) { $diText = $diText.Replace("{Theme_$key}", $colors[$key]) }
                                $diReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($diText))
                                $diWin = [System.Windows.Markup.XamlReader]::Load($diReader)
                                $diWin.Owner = $Window
                                $diWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $diWin.DragMove() })
                                $diWin.FindName("btnDiOk").Add_Click({ $diWin.Close() }.GetNewClosure())
                                if ($Window) { Show-CenteredOnOwner -ChildWindow $diWin -OwnerWindow $Window }
                                else { $diWin.WindowStartupLocation = "CenterScreen" }
                                Add-AppLog -Event "Query" -Username "System" -Details "Device info retrieved for $computer." -Config $Config -State $State -Status "Success"
                                $diWin.ShowDialog() | Out-Null
                            } elseif ($info -and -not $info.Success) {
                                Show-AppMessageBox -Message "Failed to retrieve device info for $computer.`n`nError: $($info.ErrorMessage)" -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                            } else {
                                Show-AppMessageBox -Message "No data was returned for $computer." -Title "No Data" -IconType "Warning" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                            }
                        } else {
                            Stop-Job $job -ErrorAction SilentlyContinue
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                            $reason = if ($job.ChildJobs -and $job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Connection timed out." }
                            Show-AppMessageBox -Message "Failed to query device info for $computer.`n`nError: $reason" -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                        }
                    }
                }.GetNewClosure()
                $timer.Add_Tick($timerTick)

                $btnCancelJob = $loadWin.FindName("btnCancelJob")
                if ($btnCancelJob) {
                    $btnCancelJob.Add_Click({
                        $timer.Stop()
                        try {
                            if ($job -and $job.State -eq 'Running') { Stop-Job $job -Force -ErrorAction SilentlyContinue }
                            Remove-Job $job -Force -ErrorAction SilentlyContinue
                        } catch {}
                        $loadWin.Close()
                    }.GetNewClosure())
                }

                $timer.Start()
            }
        }.GetNewClosure())
    }

    # --- System Manager ---
    if ($ctxViewProcesses) {
        $ctxViewProcesses.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $comp = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $procWin = Load-XamlWindow -XamlPath (Join-Path $AppRoot "UI\Windows\ProcessManager.xaml") -ThemeColors $colors
                $procWin.Owner = $Window
                $procWin.Title = "System Manager - $comp"
                
                $lblHeaderTitle = $procWin.FindName("lblHeaderTitle"); if ($lblHeaderTitle) { $lblHeaderTitle.Text = "Managing $comp" }
                
                $tabControlMain = $procWin.FindName("tabControlMain")
                $lvSoftware = $procWin.FindName("lvSoftware")
                $lblDiskSpace = $procWin.FindName("lblDiskSpace")
                $lvProcesses = $procWin.FindName("lvProcesses")
                $lvServices = $procWin.FindName("lvServices")
                $lvProfiles = $procWin.FindName("lvProfiles")
                $lvDevices = $procWin.FindName("lvDevices")
                $lvEvents = $procWin.FindName("lvEvents")
                $lblProcStatus = $procWin.FindName("lblProcStatus")
                $lblUptime = $procWin.FindName("lblUptime")
                
                $tab0Content = $tabControlMain.Items[0].Content
                $lblDeviceInfoStatus = $tab0Content.FindName("lblDeviceInfoStatus")
                $btnRefreshDeviceInfo = $tab0Content.FindName("btnRefreshDeviceInfo")
                
                $btnStartProcess = $procWin.FindName("btnStartProcess")
                $btnRefreshProcs = $procWin.FindName("btnRefreshProcs")
                $btnCloseProcs = $procWin.FindName("btnCloseProcs")
                $chkAutoRefreshProcs = $procWin.FindName("chkAutoRefreshProcs")
                
                $ctxKillProcess = $procWin.FindName("ctxKillProcess")
                $ctxStartService = $procWin.FindName("ctxStartService")
                $ctxStopService = $procWin.FindName("ctxStopService")
                $ctxRestartService = $procWin.FindName("ctxRestartService")
                $ctxDeleteProfile = $procWin.FindName("ctxDeleteProfile")
                $ctxUninstallSoftware = $procWin.FindName("ctxUninstallSoftware")
                $ctxEnableDevice = $procWin.FindName("ctxEnableDevice")
                $ctxDisableDevice = $procWin.FindName("ctxDisableDevice")
                
                $State.SoftLastSortCol = $null; $State.SoftSortDesc = $false
                $State.ProcLastSortCol = $null; $State.ProcSortDesc = $false
                $State.SvcLastSortCol = $null; $State.SvcSortDesc = $false
                $State.ProfLastSortCol = $null; $State.ProfSortDesc = $false
                $State.DevLastSortCol = $null;  $State.DevSortDesc = $false
                $State.EvtLastSortCol = $null;  $State.EvtSortDesc = $false
                $State.IsProcRefreshing = $false; $State.LastTabIndex = 0

                $DoRefresh = {
                    if ($State.IsProcRefreshing) { return }
                    $State.IsProcRefreshing = $true
                    $idx = if ($tabControlMain) { $tabControlMain.SelectedIndex } else { 0 }
                    
                    if ($lblProcStatus) { 
                        if ($idx -eq 0) { $lblProcStatus.Text = "Refreshing hardware device info..." }
                        elseif ($idx -eq 1) { $lblProcStatus.Text = "Refreshing software inventory..." }
                        elseif ($idx -eq 2) { $lblProcStatus.Text = "Refreshing processes..." }
                        elseif ($idx -eq 3) { $lblProcStatus.Text = "Refreshing services..." }
                        elseif ($idx -eq 4) { $lblProcStatus.Text = "Refreshing user profiles..." }
                        elseif ($idx -eq 5) { $lblProcStatus.Text = "Refreshing hardware devices..." }
                        elseif ($idx -eq 6) { $lblProcStatus.Text = "Querying recent warnings and errors..." }
                    }
                    
                    if ($btnRefreshProcs) { $btnRefreshProcs.IsEnabled = $false }
                    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                    
                    $frame = New-Object System.Windows.Threading.DispatcherFrame
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
                    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
                    
                    try {
                        if ($idx -eq 0) {
                            $tab0Content = $tabControlMain.Items[0].Content
                            $lblDIStatus = $tab0Content.FindName("lblDeviceInfoStatus")
                            if ($lblDIStatus) { $lblDIStatus.Text = "Refreshing hardware information for $comp..." }
                            $info = Get-RemoteDeviceInfo -ComputerName $comp
                            if ($info) {
                                $tab0Content.FindName("txtDI_ComputerName").Text        = if ($info.ComputerName)        { $info.ComputerName }        else { "-" }
                                $tab0Content.FindName("txtDI_Manufacturer").Text        = if ($info.Manufacturer)        { $info.Manufacturer }        else { "-" }
                                $tab0Content.FindName("txtDI_Model").Text               = if ($info.Model)               { $info.Model }               else { "-" }
                                $tab0Content.FindName("txtDI_SystemFamily").Text        = if ($info.SystemFamily)        { $info.SystemFamily }        else { "-" }
                                $tab0Content.FindName("txtDI_SerialNumber").Text        = if ($info.SerialNumber)        { $info.SerialNumber }        else { "-" }
                                $tab0Content.FindName("txtDI_BIOSVersion").Text         = if ($info.BIOSVersion)         { $info.BIOSVersion }         else { "-" }
                                $tab0Content.FindName("txtDI_BatteryStatus").Text       = if ($info.BatteryStatus)       { $info.BatteryStatus }       else { "-" }
                                $tab0Content.FindName("txtDI_AdminPasswordStatus").Text = if ($null -ne $info.AdminPasswordStatus) { "$($info.AdminPasswordStatus)" } else { "-" }
                                $tab0Content.FindName("txtDI_QueryTime").Text           = $info.QueryTime
                                if ($lblDIStatus) { $lblDIStatus.Text = "Hardware information loaded for $comp." }
                            } else {
                                if ($lblDIStatus) { $lblDIStatus.Text = "No data returned. Computer may be offline or access denied." }
                            }
                        } elseif ($idx -eq 1) {
                            $disk = Get-RemoteDiskSpace -ComputerName $comp
                            if ($lblDiskSpace -and $disk) {
                                $sizeGB = [math]::Round($disk.Size / 1GB, 1); $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
                                $lblDiskSpace.Text = "C:\ Drive: $freeGB GB Free out of $sizeGB GB ($([math]::Round(($freeGB / $sizeGB) * 100, 1))% Free)"
                            }
                            $rawSoft = Get-HDRemoteSoftware -ComputerName $comp; $resSoft = @()
                            if ($rawSoft) { 
                                foreach ($r in $rawSoft) { 
                                    $resSoft += [PSCustomObject]@{ 
                                        DisplayName = $r.Name
                                        DisplayVersion = $r.Version
                                        Publisher = $r.Type
                                        InstallDate = "-"
                                        UninstallString = $r.Identifier
                                        QuietUninstallString = $r.Identifier
                                        Type = $r.Type
                                    } 
                                } 
                            }
                            if ($lvSoftware) { if ($State.SoftLastSortCol) { $resSoft = $resSoft | Sort-Object -Property $State.SoftLastSortCol -Descending:$State.SoftSortDesc }; $lvSoftware.ItemsSource = $resSoft }
                        } elseif ($idx -eq 2) {
                            $rawProcs = Get-RemoteProcesses -ComputerName $comp; $resProcs = @()
                            if ($rawProcs) { foreach ($r in $rawProcs) { $resProcs += [PSCustomObject]@{ Name = $r.Name; Id = $r.Id; CPU = $r.CPU; MemMB = $r.MemMB; Description = $r.Description } } }
                            if ($lvProcesses) { if ($State.ProcLastSortCol) { $resProcs = $resProcs | Sort-Object -Property $State.ProcLastSortCol -Descending:$State.ProcSortDesc }; $lvProcesses.ItemsSource = $resProcs }
                        } elseif ($idx -eq 3) {
                            $rawSvcs = Get-RemoteServices -ComputerName $comp; $resSvcs = @()
                            if ($rawSvcs) { foreach ($s in $rawSvcs) { $resSvcs += [PSCustomObject]@{ Name = $s.Name; DisplayName = $s.DisplayName; State = $s.State; StartMode = $s.StartMode } } }
                            if ($lvServices) { if ($State.SvcLastSortCol) { $resSvcs = $resSvcs | Sort-Object -Property $State.SvcLastSortCol -Descending:$State.SvcSortDesc } else { $resSvcs = $resSvcs | Sort-Object -Property Name }; $lvServices.ItemsSource = $resSvcs }
                        } elseif ($idx -eq 4) {
                            $rawProfs = Get-RemoteUserProfiles -ComputerName $comp; $resProfs = @()
                            if ($rawProfs) { foreach ($p in $rawProfs) { $resProfs += [PSCustomObject]@{ LocalPath = $p.LocalPath; LastUseTime = $p.LastUseTime; Loaded = $p.Loaded; SID = $p.SID } } }
                            if ($lvProfiles) { if ($State.ProfLastSortCol) { $resProfs = $resProfs | Sort-Object -Property $State.ProfLastSortCol -Descending:$State.ProfSortDesc }; $lvProfiles.ItemsSource = $resProfs }
                        } elseif ($idx -eq 5) {
                            $rawDevs = Get-RemoteDevices -ComputerName $comp; $resDevs = @()
                            if ($rawDevs) { foreach ($d in $rawDevs) { $resDevs += [PSCustomObject]@{ FriendlyName = $d.FriendlyName; Class = $d.Class; Status = $d.Status; Manufacturer = $d.Manufacturer; InstanceId = $d.InstanceId } } }
                            if ($lvDevices) { if ($State.DevLastSortCol) { $resDevs = $resDevs | Sort-Object -Property $State.DevLastSortCol -Descending:$State.DevSortDesc } else { $resDevs = $resDevs | Sort-Object -Property Class, FriendlyName }; $lvDevices.ItemsSource = $resDevs }
                        } elseif ($idx -eq 6) {
                            $rawEvts = Get-RemoteEventLogs -ComputerName $comp; $resEvts = @()
                            if ($rawEvts) { foreach ($e in $rawEvts) { $resEvts += [PSCustomObject]@{ TimeCreated = $e.TimeCreated; Level = $e.LevelDisplayName; Id = $e.Id; Source = $e.ProviderName; Message = if ($e.Message) { $e.Message -replace "`r", "" -replace "`n", "  " } else { "" } } } }
                            if ($lvEvents) { if ($State.EvtLastSortCol) { $resEvts = $resEvts | Sort-Object -Property $State.EvtLastSortCol -Descending:$State.EvtSortDesc }; $lvEvents.ItemsSource = $resEvts }
                        }

                        try {
                            $upJob = Start-Job -ScriptBlock {
                                param($c)
                                try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $c -ErrorAction Stop; $lastBoot = $os.LastBootUpTime; $uptime = (Get-Date) - $lastBoot; return [PSCustomObject]@{ Success=$true; Days=$uptime.Days; Hours=$uptime.Hours; Minutes=$uptime.Minutes } } 
                                catch { return [PSCustomObject]@{ Success=$false; ErrorMessage=$_.Exception.Message } }
                            } -ArgumentList $comp
                            
                            $tCount = 40; while ($upJob.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
                            if ($upJob.State -eq 'Completed') {
                                $upData = Receive-Job $upJob -ErrorAction SilentlyContinue
                                if ($upData -and $upData.Success) { if ($lblUptime) { $lblUptime.Text = "Uptime: $($upData.Days) days, $($upData.Hours) hours, $($upData.Minutes) minutes" } } 
                                else { $msg = if ($upData -and $upData.ErrorMessage) { $upData.ErrorMessage } else { "Unknown Error" }; if ($lblUptime) { $lblUptime.Text = "Uptime: Unavailable ($msg)" } }
                            } else { Stop-Job $upJob -ErrorAction SilentlyContinue; if ($lblUptime) { $lblUptime.Text = "Uptime: Timeout (No Response)" } }
                            Remove-Job $upJob -Force -ErrorAction SilentlyContinue
                        } catch { if ($lblUptime) { $lblUptime.Text = "Uptime: Error - $($_.Exception.Message)" } }

                        if ($lblProcStatus) { $autoMode = if ($chkAutoRefreshProcs -and $chkAutoRefreshProcs.IsChecked) { " (Auto-refresh: 15s)" } else { "" }; $lblProcStatus.Text = "Last updated: $(Get-Date -Format 'HH:mm:ss')$autoMode" }
                    } catch { if ($lblProcStatus) { $lblProcStatus.Text = "Error: $($_.Exception.Message)" } }
                    
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($btnRefreshProcs) { $btnRefreshProcs.IsEnabled = $true }
                    $State.IsProcRefreshing = $false
                }.GetNewClosure()
                
                if ($btnRefreshProcs) { $btnRefreshProcs.Add_Click($DoRefresh) }
                if ($btnCloseProcs) { $btnCloseProcs.Add_Click({ $procWin.Close() }.GetNewClosure()) }
                if ($tabControlMain) { $tabControlMain.Add_SelectionChanged({ param($sender, $e) if ($e.OriginalSource -eq $tabControlMain -and $State.LastTabIndex -ne $tabControlMain.SelectedIndex) { $State.LastTabIndex = $tabControlMain.SelectedIndex; & $DoRefresh } }.GetNewClosure()) }

                $ListSortAction = {
                    param($sender, $e)
                    $source = $e.OriginalSource
                    while ($source -and -not ($source -is [System.Windows.Controls.GridViewColumnHeader])) { if ($source -is [System.Windows.FrameworkElement]) { $source = $source.Parent } else { break } }
                    if ($source -and ($source -is [System.Windows.Controls.GridViewColumnHeader]) -and $source.Role -ne "Padding") {
                        $column = $source.Column
                        if ($column -and $column.DisplayMemberBinding) {
                            $sortBy = $column.DisplayMemberBinding.Path.Path
                            $isDesc = $false
                            if ($sender.Name -eq "lvSoftware") { if ($State.SoftLastSortCol -eq $sortBy) { $State.SoftSortDesc = -not $State.SoftSortDesc } else { $State.SoftSortDesc = $false; $State.SoftLastSortCol = $sortBy }; $isDesc = $State.SoftSortDesc }
                            elseif ($sender.Name -eq "lvProcesses") { if ($State.ProcLastSortCol -eq $sortBy) { $State.ProcSortDesc = -not $State.ProcSortDesc } else { $State.ProcSortDesc = $false; $State.ProcLastSortCol = $sortBy }; $isDesc = $State.ProcSortDesc }
                            elseif ($sender.Name -eq "lvServices") { if ($State.SvcLastSortCol -eq $sortBy) { $State.SvcSortDesc = -not $State.SvcSortDesc } else { $State.SvcSortDesc = $false; $State.SvcLastSortCol = $sortBy }; $isDesc = $State.SvcSortDesc }
                            elseif ($sender.Name -eq "lvProfiles") { if ($State.ProfLastSortCol -eq $sortBy) { $State.ProfSortDesc = -not $State.ProfSortDesc } else { $State.ProfSortDesc = $false; $State.ProfLastSortCol = $sortBy }; $isDesc = $State.ProfSortDesc }
                            elseif ($sender.Name -eq "lvDevices") { if ($State.DevLastSortCol -eq $sortBy) { $State.DevSortDesc = -not $State.DevSortDesc } else { $State.DevSortDesc = $false; $State.DevLastSortCol = $sortBy }; $isDesc = $State.DevSortDesc }
                            elseif ($sender.Name -eq "lvEvents") { if ($State.EvtLastSortCol -eq $sortBy) { $State.EvtSortDesc = -not $State.EvtSortDesc } else { $State.EvtSortDesc = $false; $State.EvtLastSortCol = $sortBy }; $isDesc = $State.EvtSortDesc }

                            if ($sender.ItemsSource) { $items = @($sender.ItemsSource); if ($items.Count -gt 0) { $sorted = $items | Sort-Object -Property $sortBy -Descending:$isDesc; $sender.ItemsSource = @($sorted) } }
                        }
                    }
                }.GetNewClosure()
                
                if ($lvSoftware) { $lvSoftware.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvProcesses) { $lvProcesses.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvServices) { $lvServices.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvProfiles) { $lvProfiles.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvDevices) { $lvDevices.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvEvents) { $lvEvents.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }

                if ($btnRefreshDeviceInfo) {
                    $btnRefreshDeviceInfo.Add_Click({
                        if ($lblDeviceInfoStatus) { $lblDeviceInfoStatus.Text = "Refreshing hardware information for $comp..." }
                        $action = [System.Action]{
                            try {
                                $info = Get-RemoteDeviceInfo -ComputerName $comp
                                $tab0Content.Dispatcher.Invoke({
                                    if ($info) {
                                        $tab0Content.FindName("txtDI_ComputerName").Text        = if ($info.ComputerName)        { $info.ComputerName }        else { "-" }
                                        $tab0Content.FindName("txtDI_Manufacturer").Text        = if ($info.Manufacturer)        { $info.Manufacturer }        else { "-" }
                                        $tab0Content.FindName("txtDI_Model").Text               = if ($info.Model)               { $info.Model }               else { "-" }
                                        $tab0Content.FindName("txtDI_SystemFamily").Text        = if ($info.SystemFamily)        { $info.SystemFamily }        else { "-" }
                                        $tab0Content.FindName("txtDI_SerialNumber").Text        = if ($info.SerialNumber)        { $info.SerialNumber }        else { "-" }
                                        $tab0Content.FindName("txtDI_BIOSVersion").Text         = if ($info.BIOSVersion)         { $info.BIOSVersion }         else { "-" }
                                        $tab0Content.FindName("txtDI_BatteryStatus").Text       = if ($info.BatteryStatus)       { $info.BatteryStatus }       else { "-" }
                                        $tab0Content.FindName("txtDI_AdminPasswordStatus").Text = if ($null -ne $info.AdminPasswordStatus) { "$($info.AdminPasswordStatus)" } else { "-" }
                                        $tab0Content.FindName("txtDI_QueryTime").Text           = $info.QueryTime
                                        if ($lblDeviceInfoStatus) { $lblDeviceInfoStatus.Text = "Hardware information loaded for $comp." }
                                    } else {
                                        if ($lblDeviceInfoStatus) { $lblDeviceInfoStatus.Text = "No data returned. Computer may be offline or access denied." }
                                    }
                                })
                            } catch {
                                $err = $_.Exception.Message
                                $tab0Content.Dispatcher.Invoke({ if ($lblDeviceInfoStatus) { $lblDeviceInfoStatus.Text = "Failed to load: $err" } })
                            }
                        }
                        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $action) | Out-Null
                    }.GetNewClosure())
                }

                if ($ctxUninstallSoftware) {
                    $ctxUninstallSoftware.Add_Click({
                        if ($lvSoftware -and $lvSoftware.SelectedItem) {
                            $app = $lvSoftware.SelectedItem; $dispName = $app.DisplayName
                            $conf = Show-AppMessageBox -Message "Are you sure you want to silently uninstall '$dispName' from $comp?" -Title "Confirm Uninstall" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors
                            if ($conf -eq "Yes") {
                                try { 
                                    $success = Uninstall-HDRemoteSoftware -ComputerName $comp -Identifier $app.QuietUninstallString -Type $app.Type
                                    if ($success) {
                                        Show-AppMessageBox -Message "Uninstall command triggered successfully." -Title "Success" -ThemeColors $colors | Out-Null
                                        & $DoRefresh 
                                    } else {
                                        Show-AppMessageBox -Message "Uninstall failed. Check console or event logs." -Title "Error" -IconType "Error" -ThemeColors $colors | Out-Null
                                    }
                                }
                                catch { Show-AppMessageBox -Message "Uninstall failed:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors | Out-Null }
                            }
                        }
                    }.GetNewClosure())
                }
                
                if ($btnStartProcess) {
                    $btnStartProcess.Add_Click({
                        $inputXaml = @"
                        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Run Task" Width="380" SizeToContent="Height" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                            <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                                <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                                <Grid>
                                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                    <Border Grid.Row="0" Background="Transparent" Padding="16,16,16,8"><TextBlock Text="Run New Task on $($comp)" FontSize="15" FontWeight="SemiBold" Foreground="{Theme_Fg}"/></Border>
                                    <StackPanel Grid.Row="1" Margin="16,8,16,16"><TextBlock Text="Executable or Command Line:" FontSize="12" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><TextBox x:Name="txtCmd" Height="30" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Padding="6,4" VerticalContentAlignment="Center"/></StackPanel>
                                    <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}"><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button x:Name="btnOk" Content="Run" Width="80" Height="28" Margin="0,0,8,0" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsDefault="True"/><Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" Background="{Theme_Bg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" IsCancel="True"/></StackPanel></Border>
                                </Grid>
                            </Border>
                        </Window>
"@
                        foreach ($key in $colors.Keys) { $inputXaml = $inputXaml.Replace("{Theme_$key}", $colors[$key]) }
                        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($inputXaml)); $inpWin = [System.Windows.Markup.XamlReader]::Load($reader); $inpWin.Owner = $procWin
                        
                        $btnOk = $inpWin.FindName("btnOk"); $btnCancel = $inpWin.FindName("btnCancel"); $txtCmd = $inpWin.FindName("txtCmd")
                        if ($btnCancel) { $btnCancel.Add_Click({ $inpWin.Close() }.GetNewClosure()) }
                        if ($btnOk) {
                            $btnOk.Add_Click({
                                $cmd = $txtCmd.Text; $inpWin.Close()
                                if ([string]::IsNullOrWhiteSpace($cmd)) { return }
                                try { Start-RemoteProcess -ComputerName $comp -CommandLine $cmd; Add-AppLog -Event "Task Started" -Username "System" -Details "Executed '$cmd' on $comp." -Config $Config -State $State -Status "Success"; & $DoRefresh } 
                                catch { Show-AppMessageBox -Message "Failed to start task:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $procWin -ThemeColors $colors }
                            }.GetNewClosure())
                        }
                        $inpWin.Show()
                    }.GetNewClosure())
                }

                if ($ctxKillProcess) { $ctxKillProcess.Add_Click({ if ($lvProcesses -and $lvProcesses.SelectedItem) { $pidToKill = $lvProcesses.SelectedItem.Id; $procName = $lvProcesses.SelectedItem.Name; $conf = Show-AppMessageBox -Message "Kill process '$procName' (PID: $pidToKill) on $comp?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { try { Stop-RemoteProcess -ComputerName $comp -ProcessId $pidToKill; Show-AppMessageBox -Message "Process killed." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { Show-AppMessageBox -Message "Failed to kill process: $($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()) }
                if ($ctxDeleteProfile) { $ctxDeleteProfile.Add_Click({ if ($lvProfiles -and $lvProfiles.SelectedItem) { $pPath = $lvProfiles.SelectedItem.LocalPath; $pSID = $lvProfiles.SelectedItem.SID; $conf = Show-AppMessageBox -Message "Permanently delete user profile '$pPath' on $comp?`n`nThis cannot be undone." -Title "Confirm Delete" -ButtonType "YesNo" -IconType "Error" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { try { Remove-RemoteUserProfile -ComputerName $comp -SID $pSID; Show-AppMessageBox -Message "Profile deleted." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { Show-AppMessageBox -Message "Failed to delete profile: $($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()) }

                $ExecuteServiceAction = { param($ActionName) if ($lvServices -and $lvServices.SelectedItem) { $svcName = $lvServices.SelectedItem.Name; $dispName = $lvServices.SelectedItem.DisplayName; $conf = Show-AppMessageBox -Message "Are you sure you want to $ActionName the service '$dispName' on $comp?" -Title "Confirm $ActionName" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait; try { Invoke-RemoteServiceAction -ComputerName $comp -ServiceName $svcName -Action $ActionName; [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Service command sent successfully." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Service error:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()
                if ($ctxStartService) { $ctxStartService.Add_Click({ & $ExecuteServiceAction "Start" }.GetNewClosure()) }
                if ($ctxStopService) { $ctxStopService.Add_Click({ & $ExecuteServiceAction "Stop" }.GetNewClosure()) }
                if ($ctxRestartService) { $ctxRestartService.Add_Click({ & $ExecuteServiceAction "Restart" }.GetNewClosure()) }
                
                $ExecuteDeviceAction = { param($ActionName) if ($lvDevices -and $lvDevices.SelectedItem) { $devName = $lvDevices.SelectedItem.FriendlyName; $devId = $lvDevices.SelectedItem.InstanceId; $conf = Show-AppMessageBox -Message "Are you sure you want to $ActionName the device '$devName' on $comp?" -Title "Confirm $ActionName" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait; try { Set-RemoteDeviceState -ComputerName $comp -InstanceId $devId -Action $ActionName; [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Device $ActionName command sent successfully." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Device error:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()
                if ($ctxEnableDevice) { $ctxEnableDevice.Add_Click({ & $ExecuteDeviceAction "Enable" }.GetNewClosure()) }
                if ($ctxDisableDevice) { $ctxDisableDevice.Add_Click({ & $ExecuteDeviceAction "Disable" }.GetNewClosure()) }

                & $DoRefresh
                $procAutoTimer = New-Object System.Windows.Threading.DispatcherTimer
                $procAutoTimer.Interval = [TimeSpan]::FromSeconds(15)
                $procAutoTimer.Add_Tick({ & $DoRefresh }.GetNewClosure())
                
                if ($chkAutoRefreshProcs) {
                    $chkAutoRefreshProcs.Add_Checked({ $procAutoTimer.Start(); & $DoRefresh }.GetNewClosure())
                    $chkAutoRefreshProcs.Add_Unchecked({ $procAutoTimer.Stop() }.GetNewClosure())
                    if ($chkAutoRefreshProcs.IsChecked -eq $true) { $procAutoTimer.Start() }
                }
                
                $procWin.Add_Closed({ if ($procAutoTimer) { $procAutoTimer.Stop() } }.GetNewClosure())

                $procWin.Add_KeyDown({
                    param($sender, $e)
                    if ($e.Key -eq 'F5') { $e.Handled = $true; & $DoRefresh }
                    if ($e.Key -eq 'Escape') { $e.Handled = $true; $procWin.Close() }
                    if ($e.Key -eq 'Delete') {
                        $lvProcs = $procWin.FindName("lvProcesses")
                        if ($lvProcs -and $lvProcs.SelectedItem) {
                            $proc = $lvProcs.SelectedItem
                            $conf = Show-AppMessageBox -Message "End process '$($proc.Name)' (PID $($proc.Id)) on $comp?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors (Get-FluentThemeColors $State)
                            if ($conf -eq "Yes") {
                                try {
                                    Invoke-Command -ComputerName $comp -ScriptBlock { param($pid) Stop-Process -Id $pid -Force } -ArgumentList $proc.Id
                                    Add-AppLog -Event "Process Killed" -Username "System" -Details "Ended '$($proc.Name)' (PID $($proc.Id)) on $comp." -Config $Config -State $State -Status "Warning" -Color "Orange"
                                    & $DoRefresh
                                } catch {
                                    Show-AppMessageBox -Message "Failed to end process:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $procWin -ThemeColors (Get-FluentThemeColors $State) | Out-Null
                                }
                            }
                            $e.Handled = $true
                        }
                    }
                }.GetNewClosure())

                Show-CenteredOnOwner -ChildWindow $procWin -OwnerWindow $Window
                $procWin.Show()
            }
        }.GetNewClosure())
    }

    if ($ctxPrinterMenu) {
        $ctxPrinterMenu.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $targetPC = $lvData.SelectedItem.Name
                $pmScript = Join-Path $AppRoot "PrinterManager.ps1"
                if (-not (Test-Path $pmScript)) { Show-AppMessageBox -Message "Script not found at:`n$pmScript" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); return }
                Add-AppLog -Event "Printer Management" -Username "System" -Details "Launching Printer Manager for $targetPC..." -Config $Config -State $State -Status "Info"
                try { Start-Process -FilePath "powershell.exe" -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$pmScript`" -ComputerName `"$targetPC`" -Theme `"$($State.CurrentTheme)`"" -WindowStyle Hidden } 
                catch { Show-AppMessageBox -Message "Launch Failed:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
            }
        }.GetNewClosure())
    }    

    if ($ctxActiveUsers) {
        $ctxActiveUsers.Add_SubmenuOpened({
            param($sender, $e)
            if (-not $lvData.SelectedItem -or $lvData.SelectedItem.Type -ne "Computer") { return }
            $comp = $lvData.SelectedItem.Name
            if ($sender.Tag -eq $comp) { return }

            $sender.Items.Clear()
            $loadingItem = New-Object System.Windows.Controls.MenuItem
            $loadingItem.Header = "Querying active sessions..."
            $loadingItem.IsEnabled = $false
            [void]($sender.Items.Add($loadingItem))

            $job = Start-Job -ScriptBlock {
                param($c)
                $out = quser /server:$c 2>&1
                return [PSCustomObject]@{ ExitCode = $LASTEXITCODE; Lines = $out }
            } -ArgumentList $comp

            $startTime = Get-Date
            $menuRef = $sender

            $pollTimer = New-Object System.Windows.Threading.DispatcherTimer
            $pollTimer.Interval = [TimeSpan]::FromMilliseconds(400)

            $pollTick = {
                $elapsed = ((Get-Date) - $startTime).TotalSeconds
                if ($job.State -ne 'Running' -or $elapsed -ge 12) {
                    $pollTimer.Stop()
                    $menuRef.Items.Clear()

                    if ($job.State -eq 'Completed') {
                        $result = Receive-Job $job -ErrorAction SilentlyContinue
                        Remove-Job $job -Force -ErrorAction SilentlyContinue

                        $sessions = @()
                        if ($result -and $result.ExitCode -eq 0 -and $result.Lines -and $result.Lines.Count -gt 1) {
                            $header = $result.Lines[0]
                            $colU = $header.ToUpper().IndexOf("USERNAME");   if ($colU -lt 0) { $colU = 1 }
                            $colS = $header.ToUpper().IndexOf("SESSIONNAME"); if ($colS -lt 0) { $colS = 23 }
                            $colI = $header.ToUpper().IndexOf("ID");          if ($colI -lt 0) { $colI = 42 }
                            $colSt = $header.ToUpper().IndexOf("STATE");      if ($colSt -lt 0) { $colSt = 48 }

                            for ($i = 1; $i -lt $result.Lines.Count; $i++) {
                                $ln = ($result.Lines[$i] -replace '^>', ' ').PadRight($colSt + 10)
                                try {
                                    $uN  = $ln.Substring($colU,  [Math]::Max(0, $colS  - $colU)).Trim()
                                    $sId = $ln.Substring($colI,  [Math]::Max(0, $colSt - $colI)).Trim()
                                    $stL = [Math]::Min(8, $ln.Length - $colSt)
                                    $st  = if ($stL -gt 0) { $ln.Substring($colSt, $stL).Trim() } else { "" }
                                    if ($uN -and $sId) { $sessions += [PSCustomObject]@{ Username=$uN; SessionId=$sId; State=$st } }
                                } catch {}
                            }
                        }

                        if ($sessions.Count -eq 0) {
                            $noUsers = New-Object System.Windows.Controls.MenuItem
                            $noUsers.Header = "No active users"
                            $noUsers.IsEnabled = $false
                            [void]($menuRef.Items.Add($noUsers))
                        } else {
                            foreach ($s in $sessions) {
                                $uItem = New-Object System.Windows.Controls.MenuItem
                                $uItem.Header = "$($s.Username) (ID: $($s.SessionId), $($s.State))"
                                $lItem = New-Object System.Windows.Controls.MenuItem
                                $lItem.Header = "Logoff $($s.Username)"
                                $lItem.Foreground = [System.Windows.Media.Brushes]::Red
                                $closure = {
                                    param($bU, $bI, $bC)
                                    $action = {
                                        $conf = Show-AppMessageBox -Message "Logoff $bU from $bC?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                                        if ($conf -eq "Yes") {
                                            try { Stop-RemoteUserSession -ComputerName $bC -SessionId $bI; Show-AppMessageBox -Message "Session logged off successfully." -Title "Success" -ThemeColors (Get-FluentThemeColors $State) }
                                            catch { Show-AppMessageBox -Message "Failed to logoff session:`n$_" -Title "Error" -IconType "Error" -ThemeColors (Get-FluentThemeColors $State) }
                                        }
                                    }.GetNewClosure()
                                    return $action
                                }
                                $lItem.Add_Click((& $closure $s.Username $s.SessionId $comp))
                                [void]($uItem.Items.Add($lItem))
                                [void]($menuRef.Items.Add($uItem))
                            }
                        }
                        $menuRef.Tag = $comp
                    } else {
                        Stop-Job $job -ErrorAction SilentlyContinue
                        Remove-Job $job -Force -ErrorAction SilentlyContinue
                        $errItem = New-Object System.Windows.Controls.MenuItem
                        $errItem.Header = "Query timed out or failed"
                        $errItem.IsEnabled = $false
                        [void]($menuRef.Items.Add($errItem))
                    }
                }
            }.GetNewClosure()

            $pollTimer.Add_Tick($pollTick)
            $pollTimer.Start()
        }.GetNewClosure())
    }

    if ($ctxPSSession) {
        $ctxPSSession.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                Open-RemotePowerShell -ComputerName $lvData.SelectedItem.Name -Config $Config -State $State
            }
        }.GetNewClosure())
    }
}

function Open-RemotePowerShell {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        $Config,
        $State
    )
    try {
        $psCommand = "Write-Host 'Connecting to $ComputerName...' -ForegroundColor Cyan; Enter-PSSession -ComputerName '$ComputerName'"
        $encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($psCommand))
        $argList = "-WindowStyle Normal -NoExit -ExecutionPolicy Bypass -EncodedCommand $encodedCmd"
        
        Start-Process -FilePath "powershell.exe" -ArgumentList $argList -ErrorAction Stop
        
        if ($Config -and $State) { Add-AppLog -Event "Remote Session" -Username $ComputerName -Details "Opened remote PowerShell session to $ComputerName." -Config $Config -State $State -Status "Info" -Color "Blue" }
    } catch {
        $errorMessage = $_.Exception.Message
        if ($Config -and $State) { Add-AppLog -Event "Remote Session" -Username $ComputerName -Details "Failed to open PowerShell session to ${ComputerName}: $errorMessage" -Config $Config -State $State -Status "Error" -Color "Red" }
        $themeColors = $null
        try { if ($State) { $themeColors = Get-FluentThemeColors $State } } catch {}
        Show-AppMessageBox -Message "Failed to open Remote PowerShell session to ${ComputerName}.`n`nError: $errorMessage" -Title "Connection Error" -IconType "Error" -ThemeColors $themeColors | Out-Null
    }
}

Export-ModuleMember -Function *