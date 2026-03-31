# ============================================================================
# UI.Core.psm1 - Base Application UI, Theming, Search, and Refresh
# ============================================================================

function Show-AppMessageBox {
    param(
        [string]$Message,
        [string]$Title = "Information",
        [string]$ButtonType = "OK", # OK, YesNo, OKCancel
        [string]$IconType = "Information", # Information, Warning, Error, Question
        $OwnerWindow,
        $ThemeColors
    )

    # Guard: if no theme colors provided, fall back to global State so
    # {Theme_*} tokens in MessageBox.xaml are always replaced correctly.
    if (-not $ThemeColors) {
        try { $ThemeColors = Get-FluentThemeColors $global:State } catch {}
    }
    
    $xamlPath = Join-Path $PSScriptRoot "..\UI\Dialogs\MessageBox.xaml"
    if (-not (Test-Path $xamlPath)) {
        [System.Windows.MessageBox]::Show($Message, $Title)
        return "OK"
    }

    $msgWin = Load-XamlWindow -XamlPath $xamlPath -ThemeColors $ThemeColors
    
    $lblTitle = $msgWin.FindName("lblTitle")
    if ($lblTitle) { $lblTitle.Text = $Title }
    
    $txtMessageBody = $msgWin.FindName("txtMessageBody")
    if ($txtMessageBody) { $txtMessageBody.Text = $Message }
    
    $pathIcon = $msgWin.FindName("pathIcon")
    if ($pathIcon) {
        switch ($IconType) {
            "Error" { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(209, 52, 56)) }
            "Warning" { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2L1 21h22L12 2zm0 3.83L19.53 19H4.47L12 5.83zM11 10h2v5h-2v-5zm0 6h2v2h-2v-2z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255, 185, 0)) }
            "Question" { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17h-2v-2h2v2zm2.07-7.75l-.9.92C13.45 12.9 13 13.5 13 15h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26c.37-.36.59-.86.59-1.41 0-1.1-.9-2-2-2s-2 .9-2 2H8c0-2.21 1.79-4 4-4s4 1.79 4 4c0 .88-.36 1.68-.93 2.25z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212)) }
            Default { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h-2v6h2zm0-8h-2V7h2v2z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212)) }
        }
    }
    
    $msgWin.FindName("btnYes").Visibility = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }
    $msgWin.FindName("btnNo").Visibility = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }
    $msgWin.FindName("btnOk").Visibility = if ($ButtonType -in @("OK", "OKCancel")) { "Visible" } else { "Collapsed" }
    $msgWin.FindName("btnCancel").Visibility = if ($ButtonType -eq "OKCancel") { "Visible" } else { "Collapsed" }
    
    $script:msgRes = "Cancel"
    $msgWin.FindName("btnYes").Add_Click({ $script:msgRes = "Yes"; $msgWin.Close() })
    $msgWin.FindName("btnNo").Add_Click({ $script:msgRes = "No"; $msgWin.Close() })
    $msgWin.FindName("btnOk").Add_Click({ $script:msgRes = "OK"; $msgWin.Close() })
    $msgWin.FindName("btnCancel").Add_Click({ $script:msgRes = "Cancel"; $msgWin.Close() })
    $msgWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $msgWin.DragMove() })
    
    if ($OwnerWindow -and $OwnerWindow.IsVisible) { $msgWin.Owner = $OwnerWindow; $msgWin.WindowStartupLocation = "CenterOwner" }
    $msgWin.ShowDialog() | Out-Null
    return $script:msgRes
}

# ---------------------------------------------------------------------------
# Show-CenteredOnOwner
# Centers a WPF child window on the same monitor as its owner, rather than
# always defaulting to the primary monitor. Call before ShowDialog()/Show().
# ---------------------------------------------------------------------------
function Show-CenteredOnOwner {
    param(
        [Parameter(Mandatory=$true)]  $ChildWindow,
        [Parameter(Mandatory=$false)] $OwnerWindow
    )
    if (-not $OwnerWindow) {
        $ChildWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
        return
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

        # Get the screen that contains the owner window's centre point
        $ownerLeft   = $OwnerWindow.Left
        $ownerTop    = $OwnerWindow.Top
        $ownerWidth  = $OwnerWindow.ActualWidth
        $ownerHeight = $OwnerWindow.ActualHeight
        $centrePt    = [System.Drawing.Point]::new(
            [int]($ownerLeft + $ownerWidth  / 2),
            [int]($ownerTop  + $ownerHeight / 2)
        )
        $screen = [System.Windows.Forms.Screen]::FromPoint($centrePt)
        $wa     = $screen.WorkingArea

        # Force manual positioning
        $ChildWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual

        $childW = if ($ChildWindow.Width  -gt 0) { $ChildWindow.Width  } else { 450 }
        $childH = if ($ChildWindow.Height -gt 0) { $ChildWindow.Height } else { 300 }

        $ChildWindow.Left = $wa.Left + ($wa.Width  - $childW) / 2
        $ChildWindow.Top  = $wa.Top  + ($wa.Height - $childH) / 2
    } catch {
        # Fall back to CenterOwner so we never crash
        $ChildWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
        if ($OwnerWindow -and $OwnerWindow.IsVisible) { $ChildWindow.Owner = $OwnerWindow }
    }
}

function Register-CoreUIEvents {
    param($Window, $Config, $State)

    $AppRoot = Split-Path -Path $PSScriptRoot -Parent

    # --- Ensure Image Preferences are loaded from userprefs.json ---
    $prefsFile = Join-Path $env:LOCALAPPDATA "PelicanCU\HDCompanion\userprefs.json"
    if (Test-Path -LiteralPath $prefsFile) {
        try {
            $prefs = Get-Content -LiteralPath $prefsFile -Raw | ConvertFrom-Json
            if (-not $Config.UserPreferences) { 
                $Config | Add-Member -MemberType NoteProperty -Name "UserPreferences" -Value @{} -Force 
            }
            if ($null -ne $prefs.AppBackgroundImage) {
                if ($Config.UserPreferences -is [hashtable]) { $Config.UserPreferences["AppBackgroundImage"] = $prefs.AppBackgroundImage }
                else { $Config.UserPreferences | Add-Member -MemberType NoteProperty -Name "AppBackgroundImage" -Value $prefs.AppBackgroundImage -Force }
            }
        } catch {}
    }

    # Guard: ensure UIControls is a live hashtable before we try to index into it.
    if ($null -eq $State["UIControls"]) { $State["UIControls"] = @{} }

    $State["UIControls"]["txtLog"] = $Window.FindName("txtLog")
    $lvData = $Window.FindName("lvData")
    $gvData = $Window.FindName("gvData")
    $txtSearch = $Window.FindName("txtSearch")
    $btnSearch = $Window.FindName("btnSearch")
    $btnRefresh = $Window.FindName("btnRefresh")
    $btnViewLog = $Window.FindName("btnViewLog")
    $lblStatus = $Window.FindName("lblStatus")
    $btnThemeToggle = $Window.FindName("btnThemeToggle")
    $iconTheme = $Window.FindName("iconTheme")
    $cbAutoRefresh = $Window.FindName("cbAutoRefresh")
    $chkEnableEmail = $Window.FindName("chkEnableEmail")
    
    $popSearchHistory = $null
    $lbSearchHistory  = $null

    $btnDashboard = $Window.FindName("btnDashboard")
    $btnOpenDashboard = $Window.FindName("btnOpenDashboard")
    $btnModifyConfig = $Window.FindName("btnModifyConfig")
    $btnPreferences = $Window.FindName("btnPreferences")
    $btnHelp = $Window.FindName("btnHelp")
    $btnTechDocs = $Window.FindName("btnTechDocs")
    
    # --- New Window Control Assignments ---
    $btnTopClose = $Window.FindName("btnTopClose")
    $btnBottomClose = $Window.FindName("btnBottomClose")
    $mainTitleBar = $Window.FindName("TitleBar")
    
    $ctxDetails = $Window.FindName("ctxDetails")
    $overlayDetails = $Window.FindName("overlayDetails")
    $borderDetails = $Window.FindName("borderDetails")
    $txtDetailsContent = $Window.FindName("txtDetailsContent")
    $btnCloseDetails = $Window.FindName("btnCloseDetails")
    $thumbResizeDetails = $Window.FindName("thumbResizeDetails")

    if ($txtDetailsContent) {
        $ctxCopy = New-Object System.Windows.Controls.ContextMenu
        $miCopy = New-Object System.Windows.Controls.MenuItem
        $miCopy.Header = "Copy All Details to Clipboard"
        $miCopy.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($txtDetailsContent.Text)) {
                [System.Windows.Clipboard]::SetText($txtDetailsContent.Text)
                Show-AppMessageBox -Message "Details copied to clipboard." -Title "Copied" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) | Out-Null
            }
        })
        [void]($ctxCopy.Items.Add($miCopy))
        $txtDetailsContent.ContextMenu = $ctxCopy
    }

    $ApplyTheme = {
        param($TargetTheme)
        $res = $Window.Resources
        $themeData = if ($TargetTheme -eq "Light") { $Config.LightModeColors } else { $Config.DarkModeColors }
        
        function Get-ColorFromConfig ($rgbArray) {
            if ($rgbArray -is [string]) {
                try { return [System.Windows.Media.ColorConverter]::ConvertFromString($rgbArray) } catch { return [System.Windows.Media.Colors]::Transparent }
            }
            if ($rgbArray -and $rgbArray.Count -eq 3) { return [System.Windows.Media.Color]::FromRgb($rgbArray[0], $rgbArray[1], $rgbArray[2]) }
            return [System.Windows.Media.Colors]::Transparent 
        }

        $rawWinColor = Get-ColorFromConfig $themeData.Background
        $rawCardColor = Get-ColorFromConfig $themeData.Card

        # Apply Visual Settings from userprefs.json
        if ($Config.UserPreferences) {
            $Window.Opacity = if ($Config.UserPreferences.GlassEffect -eq $true) { 0.92 } else { 1.0 }
            $res["ListItemHeight"] = if ($Config.UserPreferences.Density -eq "Compact") { [double]22 } else { [double]30 }
            
            if ($Config.UserPreferences.FontSize -eq "Small") { $res["AppFontSize"] = [double]11 }
            elseif ($Config.UserPreferences.FontSize -eq "Large") { $res["AppFontSize"] = [double]15 }
            else { $res["AppFontSize"] = [double]13 }
        }

        $res["WindowBackground"] = [System.Windows.Media.SolidColorBrush]::new($rawWinColor)
        $res["CardBackground"] = [System.Windows.Media.SolidColorBrush]::new($rawCardColor)
        $res["AccentFill"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Primary))
        $res["TextPrimary"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Text))
        $res["TextSecondary"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.TextSecondary))
        $res["ControlStroke"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Secondary))
        
        if ($themeData.Hover) { $res["HoverFill"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Hover)) } 
        else { if ($TargetTheme -eq "Light") { $res["HoverFill"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(229,229,229)) } else { $res["HoverFill"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(80,80,85)) } }
        
        if ($themeData.AltRow) { $res["AltRowBg"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.AltRow)) } 
        else { if ($TargetTheme -eq "Light") { $res["AltRowBg"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(249,249,249)) } else { $res["AltRowBg"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(55,55,60)) } }

        if ($themeData.OnlineText) { $res["OnlineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.OnlineText)) }
        else { if ($TargetTheme -eq "Dark") { $res["OnlineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(150, 255, 150)) } else { $res["OnlineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 100, 0)) } }

        if ($themeData.OfflineText) { $res["OfflineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.OfflineText)) }
        else { if ($TargetTheme -eq "Dark") { $res["OfflineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255, 150, 150)) } else { $res["OfflineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(178, 34, 34)) } }
        
        $uiLog = if ($null -ne $State["UIControls"]) { $State["UIControls"]["txtLog"] } else { $null }
        if ($uiLog) { $uiLog.Foreground = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Text)) }
        
        if ($iconTheme) {
            if ($TargetTheme -eq "Light") { $iconTheme.Data = [System.Windows.Media.Geometry]::Parse("M12 3c-4.97 0-9 4.03-9 9s4.03 9 9 9 9-4.03 9-9c0-.46-.04-.92-.1-1.36-.98 1.37-2.58 2.26-4.4 2.26-3.03 0-5.5-2.47-5.5-5.5 0-1.82.89-3.42 2.26-4.4-.44-.06-.9-.1-1.36-.1z") } 
            else { $iconTheme.Data = [System.Windows.Media.Geometry]::Parse("M6.76 4.84l-1.8-1.79-1.41 1.41 1.79 1.79 1.42-1.41zM4 10.5H1v2h3v-2zm9-9.95h-2V3.5h2V.55zm7.45 3.91l-1.41-1.41-1.79 1.79 1.41 1.41 1.79-1.79zm-3.21 13.7l1.79 1.8 1.41-1.41-1.8-1.79-1.4 1.4zM20 10.5v2h3v-2h-3zm-8-5c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6-2.69-6-6-6zm-1 16.95h2V19.5h-2v2.95zm-7.45-3.91l1.41 1.41 1.79-1.8-1.41-1.41-1.79 1.8z") }
        }
        
        # --- Handle Application Background Image ---
        $res["WindowBackgroundImage"] = [System.Windows.Media.Brushes]::Transparent
        $panelAlpha = 255

        if ($Config.UserPreferences -and -not [string]::IsNullOrWhiteSpace($Config.UserPreferences.AppBackgroundImage)) {
            $imagePath = $Config.UserPreferences.AppBackgroundImage
            if (Test-Path -LiteralPath $imagePath) {
                try {
                    $stream = [System.IO.File]::OpenRead($imagePath)
                    $img = [System.Windows.Media.Imaging.BitmapImage]::new()
                    $img.BeginInit()
                    $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $img.StreamSource = $stream
                    $img.EndInit()
                    $img.Freeze() 
                    if ($null -ne $stream) { $stream.Close(); $stream.Dispose() }
                    
                    $brush = [System.Windows.Media.ImageBrush]::new($img)
                    $brush.Stretch = [System.Windows.Media.Stretch]::UniformToFill
                    $brush.Freeze()
                    
                    $res["WindowBackgroundImage"] = [System.Windows.Media.Brush]($brush.psobject.BaseObject)
                    $panelAlpha = 220 
                } catch {
                    Write-Warning "Failed to load background image: $($_.Exception.Message)"
                }
            }
        }

        $res["LeftPanelBackground"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb($panelAlpha, $rawWinColor.R, $rawWinColor.G, $rawWinColor.B))
        $res["RightPanelBackground"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb($panelAlpha, $rawCardColor.R, $rawCardColor.G, $rawCardColor.B))

        $State.CurrentTheme = $TargetTheme
    }.GetNewClosure()

    $UpdateGridColumns = {
        if (-not $gvData) { return }
        $gvData.Columns.Clear()
        $cols = @(
            @{ Header="Name"; Binding="Name"; Width=160; HasStatusIndicator=$true },
            @{ Header="Type"; Binding="Type"; Width=80 },
            @{ Header="AD Description"; Binding="Description"; Width=200 },
            @{ Header="Locked / Sys Desc"; Binding="LockedOut"; Width=150 },
            @{ Header="Last Logon"; Binding="LastLogonDate"; Width=140; Format="{0:MM/dd/yyyy}" }
        )
        foreach ($c in $cols) {
            $col = New-Object System.Windows.Controls.GridViewColumn
            $col.Width = $c.Width
            if ($c.HasStatusIndicator) {
                $cellTemplate = @"
                <DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
                    <StackPanel Orientation="Horizontal">
                        <Ellipse Width="10" Height="10" Margin="0,0,6,0" VerticalAlignment="Center">
                            <Ellipse.Style>
                                <Style TargetType="Ellipse">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding IsOnline}" Value="True"><Setter Property="Fill" Value="#4CAF50"/><Setter Property="Visibility" Value="Visible"/></DataTrigger>
                                        <DataTrigger Binding="{Binding IsOnline}" Value="False"><Setter Property="Fill" Value="#E53935"/><Setter Property="Visibility" Value="Visible"/></DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </Ellipse.Style>
                        </Ellipse>
                        <TextBlock Text="{Binding $($c.Binding)}" VerticalAlignment="Center"/>
                    </StackPanel>
                </DataTemplate>
"@
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($cellTemplate))
                $col.CellTemplate = [System.Windows.Markup.XamlReader]::Load($reader)
            } else {
                $binding = New-Object System.Windows.Data.Binding($c.Binding)
                if ($c.Format) { $binding.StringFormat = $c.Format }
                $col.DisplayMemberBinding = $binding
            }
            $headerTemplate = "<DataTemplate xmlns=`"http://schemas.microsoft.com/winfx/2006/xaml/presentation`"><TextBlock Text=`"$($c.Header)`" HorizontalAlignment=`"Left`" Margin=`"5,0,0,0`"/></DataTemplate>"
            $headerReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($headerTemplate))
            $col.HeaderTemplate = [System.Windows.Markup.XamlReader]::Load($headerReader)
            $gvData.Columns.Add($col)
        }
    }.GetNewClosure()

    $SortListAction = {
        param($sender, $e)
        $source = $e.OriginalSource
        while ($source -and -not ($source -is [System.Windows.Controls.GridViewColumnHeader])) {
            if ($source -is [System.Windows.FrameworkElement]) { $source = $source.Parent } else { break }
        }
        if ($source -and ($source -is [System.Windows.Controls.GridViewColumnHeader]) -and $source.Role -ne "Padding") {
            $column = $source.Column
            $sortBy = if ($column.DisplayMemberBinding) { $column.DisplayMemberBinding.Path.Path } else { "Name" }
            if ($sortBy) {
                if ($State.LastSortCol -eq $sortBy) { $State.SortDescending = -not $State.SortDescending } 
                else { $State.SortDescending = $false; $State.LastSortCol = $sortBy }
                if ($lvData -and $lvData.ItemsSource) {
                    $items = @($lvData.ItemsSource) 
                    if ($items.Count -gt 0) {
                         $sorted = $items | Sort-Object -Property $sortBy -Descending:$State.SortDescending
                         $lvData.ItemsSource = @($sorted)
                         if ($lblStatus) { $lblStatus.Text = "Sorted by $sortBy ($if ($State.SortDescending) { 'Descending' } else { 'Ascending' })" }
                    }
                }
            }
        }
    }.GetNewClosure()

    $PerformSearch = {
        if (-not $txtSearch) { return }
        $term = $txtSearch.Text
        if ([string]::IsNullOrWhiteSpace($term)) {
            Add-AppLog -Event "Search" -Username "System" -Details "Please enter a search term." -Config $Config -State $State -Color "Orange" -Status "Warning"
            return
        }
        
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        Add-AppLog -Event "Search" -Username "System" -Details "Searching Directory for '$term'..." -Config $Config -State $State -Status "Info"
        
        $userResults = Search-ADUsers -SearchTerm $term -Config $Config
        $compResults = Search-ADComputers -SearchTerm $term -Config $Config
        
        if ($compResults -and $compResults.Count -gt 0) {
            $compNames = @($compResults | Select-Object -ExpandProperty Name | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($compNames.Count -gt 0) {
                $job = Invoke-Command -ComputerName $compNames -ScriptBlock { 
                    try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Description } catch { "WMI Error" }
                } -AsJob -ErrorAction SilentlyContinue
                
                if ($job) {
                    $timeoutCount = 40
                    while ($job.State -eq 'Running' -and $timeoutCount -gt 0) { Start-Sleep -Milliseconds 100; $timeoutCount-- }
                    $wmiResults = Receive-Job $job -ErrorAction SilentlyContinue
                    Stop-Job $job -ErrorAction SilentlyContinue
                    Remove-Job $job -Force -ErrorAction SilentlyContinue
                    
                    foreach ($comp in $compResults) {
                        if (-not $comp.psobject.Properties['LockedOut']) { $comp | Add-Member -MemberType NoteProperty -Name "LockedOut" -Value "" -Force }
                        $match = @($wmiResults | Where-Object { $_.PSComputerName -eq $comp.Name })
                        if ($match.Count -gt 0) {
                            $descVal = $match[0]
                            if ([string]::IsNullOrWhiteSpace($descVal)) { $comp.LockedOut = "<Blank>" } else { $comp.LockedOut = [string]$descVal }
                        } else { $comp.LockedOut = "Unreachable" }
                    }
                }
            }
        }
        
        $allResults = @($userResults) + @($compResults)
        if ($lvData) { $lvData.ItemsSource = $allResults }
        if ($lblStatus) { $lblStatus.Text = "Found $($allResults.Count) objects matching '$term'. (Auto-refresh paused 10m)" }
        
        $State.IsSearchPaused = $true
        $State.RefreshTargetTime = (Get-Date).AddMinutes(10)
        if ($State.Timer -and $State.Timer.Interval.TotalSeconds -ne 1) {
            $State.Timer.Interval = [TimeSpan]::FromSeconds(1)
            if (-not $State.Timer.IsEnabled) { $State.Timer.Start() }
        }
        [System.Windows.Input.Mouse]::OverrideCursor = $null

        # --- Save to search history ---
        if (-not $State.SearchHistory) { $State.SearchHistory = @() }
        if ($term -notin $State.SearchHistory) {
            $State.SearchHistory = @($term) + @($State.SearchHistory | Select-Object -First 19)
            if ($txtSearch) {
                $txtSearch.ItemsSource = $State.SearchHistory
            }
            try {
                $prefsDir  = Join-Path $env:LOCALAPPDATA "PelicanCU\HDCompanion"
                $prefsPath = Join-Path $prefsDir "userprefs.json"
                $prefs = if (Test-Path $prefsPath) { Get-Content $prefsPath -Raw | ConvertFrom-Json } else { [PSCustomObject]@{} }
                $prefs | Add-Member -MemberType NoteProperty -Name "SearchHistory" -Value $State.SearchHistory -Force
                if (-not (Test-Path $prefsDir)) { New-Item -ItemType Directory -Path $prefsDir -Force | Out-Null }
                $prefs | ConvertTo-Json -Depth 5 | Set-Content -Path $prefsPath -Encoding UTF8
            } catch {}
        }
        if ($txtSearch -and $txtSearch.Text -ne $term) { $txtSearch.Text = $term }
    }.GetNewClosure()

    $RefreshAction = {
        if ($State.IsRefreshing) { return }
        $State.IsRefreshing = $true

        if ($State.IsSearchPaused) {
            $State.IsSearchPaused = $false
            if ($State.Timer -and -not $State.Timer.IsEnabled) { $State.Timer.Start() }
        }
        
        $State.RefreshTargetTime = (Get-Date).AddSeconds($State.RefreshIntervalSeconds)
        
        if ($btnRefresh) { $btnRefresh.Content = "Refreshing..." }
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

        if ($txtSearch) { $txtSearch.Text = "" }

        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)

        try {
            $safeLocked = Get-LockedADUsers -Config $Config
            if ($null -eq $safeLocked) { $safeLocked = @() }

            $isFirstRun = ($null -eq $State.LoggedLockoutTimes)
            if ($isFirstRun) { $State.LoggedLockoutTimes = @{} }

            $logDir   = $Config.GeneralSettings.LogDirectoryUNC
            $today    = Get-Date -Format "yyyyMMdd"
            $logFile  = Join-Path $logDir "UnlockLog_$today.csv"
            $recentLogs = $null; $logsLoaded = $false

            $currentLockedUsers = @()
            foreach ($u in $safeLocked) {
                $currentLockedUsers += $u.Name
                $newLockoutTime = if ($u.LockoutTime) { "$($u.LockoutTime)" } else { "Unknown" }
                
                if (-not $State.LoggedLockoutTimes.ContainsKey($u.Name) -or $State.LoggedLockoutTimes[$u.Name] -ne $newLockoutTime) {
                    if (-not [string]::IsNullOrWhiteSpace($newLockoutTime) -and $newLockoutTime -ne "Unknown") {
                        $isDuplicate = $false
                        if (-not $logsLoaded) {
                            if (Test-Path -LiteralPath $logFile) { try { $recentLogs = @(Import-Csv -LiteralPath $logFile -ErrorAction Stop | Select-Object -Last 100) } catch { $recentLogs = @() } } else { $recentLogs = @() }
                            $logsLoaded = $true
                        }
                        if ($recentLogs) {
                            $dup = $recentLogs | Where-Object { $_.Event -eq "Lockout Detected" -and $_.Username -eq $u.Name }
                            if ($dup) {
                                foreach ($d in $dup) {
                                    if ($d.Details -match "Time:\s*(.*)\)") {
                                        try { $lTime = [datetime]$matches[1]; $nTime = [datetime]$newLockoutTime; if ([math]::Abs(($nTime - $lTime).TotalMinutes) -lt 5) { $isDuplicate = $true; break } } catch { if ($d.Details -match [regex]::Escape($newLockoutTime)) { $isDuplicate = $true; break } }
                                    } elseif ($d.Details -match [regex]::Escape($newLockoutTime)) { $isDuplicate = $true; break }
                                }
                            }
                        }
                        $State.LoggedLockoutTimes[$u.Name] = $newLockoutTime
                        if (-not $isDuplicate) {
                            $detailsMsg = "Account lockout detected. (AD Time: $newLockoutTime)"
                            Add-AppLog -Event "Lockout Detected" -Username $u.Name -Details $detailsMsg -Config $Config -State $State -Status "Warning" -Color "Orange"
                            if ($null -ne $recentLogs) { $recentLogs += [PSCustomObject]@{ Event="Lockout Detected"; Username=$u.Name; Details=$detailsMsg; Timestamp=(Get-Date).ToString("yyyy-MM-ddTHH:mm:ss") } }
                        }
                    }
                }
            }

            if (-not $isFirstRun) {
                $clearedUsers = @()
                foreach ($trackedUser in $State.LoggedLockoutTimes.Keys) {
                    if ($trackedUser -notin $currentLockedUsers) { $clearedUsers += $trackedUser }
                }
                foreach ($cleared in $clearedUsers) {
                    $isDuplicate = $false
                    if (-not $logsLoaded) {
                        if (Test-Path -LiteralPath $logFile) { try { $recentLogs = @(Import-Csv -LiteralPath $logFile -ErrorAction Stop | Select-Object -Last 100) } catch { $recentLogs = @() } } else { $recentLogs = @() }
                        $logsLoaded = $true
                    }
                    if ($recentLogs) {
                        $dup = $recentLogs | Where-Object { $_.Event -eq "Lockout Cleared" -and $_.Username -eq $cleared }
                        foreach ($d in $dup) { try { if ([math]::Abs(((Get-Date) - [datetime]$d.Timestamp).TotalMinutes) -lt 5) { $isDuplicate = $true; break } } catch {} }
                    }
                    if (-not $isDuplicate) {
                        # Add-AppLog -Event "Lockout Cleared" -Username $cleared -Details "Account is no longer locked." -Config $Config -State $State -Status "Info" -Color "Green"
                        if ($null -ne $recentLogs) { $recentLogs += [PSCustomObject]@{ Event="Lockout Cleared"; Username=$cleared; Timestamp=(Get-Date).ToString("yyyy-MM-ddTHH:mm:ss") } }
                    }
                    $State.LoggedLockoutTimes.Remove($cleared)
                }
            }

            if ($lvData) { $lvData.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[psobject]]$safeLocked }
            if ($lblStatus) { $lblStatus.Text = "Found $($safeLocked.Count) locked accounts. (Last Update: $(Get-Date -Format 'HH:mm:ss'))" }

        } catch {
            if ($lblStatus) { $lblStatus.Text = "Refresh error: $($_.Exception.Message)" }
        } finally {
            if ($btnRefresh) { $btnRefresh.Content = "Refresh List ($($State.RefreshIntervalSeconds)s)" }
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            $State.IsRefreshing = $false
        }
    }.GetNewClosure()

    $State.Actions.RefreshData = $RefreshAction

    & $UpdateGridColumns
    if ($lvData) {
        $lvData.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$SortListAction)
        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            if ($ctxDetails) { $ctxDetails.Visibility = if ($sel) { "Visible" } else { "Collapsed" } }
        }.GetNewClosure())
    }
    
    if ($btnThemeToggle) { $btnThemeToggle.Add_Click({ $newTheme = if ($State.CurrentTheme -eq "Dark") { "Light" } else { "Dark" }; & $ApplyTheme -TargetTheme $newTheme }.GetNewClosure()) }
    if ($btnSearch) { $btnSearch.Add_Click($PerformSearch) }
    
    if ($txtSearch) {
        $txtSearch.Add_KeyDown({
            param($sender, $e)
            if ($e.Key -eq 'Enter') {
                $e.Handled = $true
                if ($txtSearch.IsDropDownOpen) { $txtSearch.IsDropDownOpen = $false }
                if (-not [string]::IsNullOrWhiteSpace($txtSearch.Text)) { & $PerformSearch }
            }
        }.GetNewClosure())

        $txtSearch.Add_SelectionChanged({
            param($sender, $e)
            if ($txtSearch.SelectedItem -and -not [string]::IsNullOrWhiteSpace($txtSearch.SelectedItem)) {
                $txtSearch.Text = $txtSearch.SelectedItem
                $txtSearch.IsDropDownOpen = $false
                & $PerformSearch
            }
        }.GetNewClosure())

        # ---------------------------------------------------------------------
        # FIX: Handling the "X" button click inside the ComboBox ItemTemplate
        # ---------------------------------------------------------------------
        $txtSearch.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
            param($sender, $e)
            $source = $e.OriginalSource
            # Check if the clicked element is a Button named 'btnRemoveHistoryItem'
            if ($source -is [System.Windows.Controls.Button] -and $source.Name -eq "btnRemoveHistoryItem") {
                # The search string is stored in the Tag property of the button
                $itemToRemove = $source.Tag
                if ($null -ne $itemToRemove) {
                    
                    # 1. Remove from In-Memory State
                    if ($State.SearchHistory) {
                        $State.SearchHistory = @($State.SearchHistory | Where-Object { $_ -ne $itemToRemove })
                    }
                    
                    # 2. Remove from UI (ComboBox Dropdown)
                    if ($txtSearch.Items.Contains($itemToRemove)) {
                        $txtSearch.Items.Remove($itemToRemove)
                    }

                    # 3. Remove from userprefs.json file
                    try {
                        $prefsDir  = Join-Path $env:LOCALAPPDATA "PelicanCU\HDCompanion"
                        $prefsPath = Join-Path $prefsDir "userprefs.json"
                        if (Test-Path -LiteralPath $prefsPath) {
                            $prefs = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
                            $prefs.SearchHistory = $State.SearchHistory
                            $prefs | ConvertTo-Json -Depth 5 | Set-Content -Path $prefsPath -Encoding UTF8
                        }
                    } catch {}
                    
                    # Stop the event from bubbling up and selecting the row we just clicked
                    $e.Handled = $true
                }
            }
        }.GetNewClosure())
    }

    $Window.Add_Loaded({
        if (-not $State.SearchHistory) { $State.SearchHistory = @() }
        try {
            $ph = Join-Path $env:LOCALAPPDATA "PelicanCU\HDCompanion\userprefs.json"
            if (Test-Path -LiteralPath $ph) {
                $savedPrefs = Get-Content -LiteralPath $ph -Raw | ConvertFrom-Json
                if ($savedPrefs.SearchHistory) {
                    $State.SearchHistory = @($savedPrefs.SearchHistory)
                }
            }
        } catch {}
        if ($txtSearch -and $State.SearchHistory.Count -gt 0) {
            $txtSearch.ItemsSource = $State.SearchHistory
        }
        & $ApplyTheme -TargetTheme $State.CurrentTheme
        if ($btnRefresh) { $btnRefresh.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
        $State.Timer.Interval = [TimeSpan]::FromSeconds(1); $State.Timer.Start()
        $cfgPath = if ($Config.LoadedConfigPath) { $Config.LoadedConfigPath } else { "Unknown" }
        Add-AppLog -Event "Config" -Username "System" -Details "Using configuration from: $cfgPath" -Config $Config -State $State -Status "Info" -Color "Blue"
    }.GetNewClosure())

    $Window.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq 'F5') {
            $e.Handled = $true
            & $RefreshAction
        }
        if ($e.Key -eq 'Escape') {
            $overlayResetKb = $Window.FindName("overlayReset")
            $overlayDetailsKb = $Window.FindName("overlayDetails")
            if ($overlayResetKb -and $overlayResetKb.Visibility -eq "Visible")    { $overlayResetKb.Visibility   = "Collapsed"; $e.Handled = $true }
            if ($overlayDetailsKb -and $overlayDetailsKb.Visibility -eq "Visible") { $overlayDetailsKb.Visibility = "Collapsed"; $e.Handled = $true }
            if ($txtSearch -and $txtSearch.IsDropDownOpen) { $txtSearch.IsDropDownOpen = $false; $e.Handled = $true }
        }
        if ($e.Key -eq 'F' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
            if ($txtSearch) { $txtSearch.Focus(); $txtSearch.SelectAll() }
            $e.Handled = $true
        }
    }.GetNewClosure())
    
    if ($btnRefresh) { 
        $btnRefresh.Add_Click($RefreshAction)
        $btnRefresh.Add_MouseEnter({ $btnRefresh.Foreground = $Window.Resources["TextPrimary"] }.GetNewClosure())
        $btnRefresh.Add_MouseLeave({ $btnRefresh.Foreground = [System.Windows.Media.Brushes]::White }.GetNewClosure())
    }

    if ($chkEnableEmail) {
        $chkEnableEmail.Add_Checked({ $Config.EmailSettings.EnableEmailNotifications = $true }.GetNewClosure())
        $chkEnableEmail.Add_Unchecked({ $Config.EmailSettings.EnableEmailNotifications = $false }.GetNewClosure())
    }

    if ($btnTopClose) { $btnTopClose.Add_Click({ $Window.Close() }.GetNewClosure()) }
    if ($btnBottomClose) { $btnBottomClose.Add_Click({ $Window.Close() }.GetNewClosure()) }
    if ($mainTitleBar) { $mainTitleBar.Add_MouseLeftButtonDown({ $Window.DragMove() }.GetNewClosure()) }

    if ($btnViewLog) {
        $btnViewLog.Add_Click({
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            $logs = Get-AppLogFiles -Config $Config
            $colors = Get-FluentThemeColors $State
            $logWin = Load-XamlWindow -XamlPath (Join-Path $AppRoot "UI\Windows\LogViewer.xaml") -ThemeColors $colors
            $logWin.Owner = $Window
            
            $lvLogs = $logWin.FindName("lvLogs")
            $txtOp = $logWin.FindName("txtFilterOperator")
            $txtUsr = $logWin.FindName("txtFilterUser")
            $dpStart = $logWin.FindName("dpStartDate")
            $dpEnd = $logWin.FindName("dpEndDate")
            
            $ApplyFilter = {
                $filtered = $logs | Where-Object {
                    $pass = $true
                    if ($txtOp.Text) { $pass = $pass -and ($_.Operator -match $txtOp.Text) }
                    if ($txtUsr.Text) { $pass = $pass -and ($_.Username -match $txtUsr.Text) }
                    if ($_.Timestamp -and $dpStart.SelectedDate) { try { if ([DateTime]$_.Timestamp -lt $dpStart.SelectedDate) { $pass = $false } } catch {} }
                    if ($_.Timestamp -and $dpEnd.SelectedDate) { try { if ([DateTime]$_.Timestamp -gt $dpEnd.SelectedDate.AddDays(1)) { $pass = $false } } catch {} }
                    return $pass
                }
                if ($lvLogs) { $lvLogs.ItemsSource = @($filtered | Sort-Object Timestamp -Descending) }
            }.GetNewClosure()
            
            & $ApplyFilter
            
            $btnFilter = $logWin.FindName("btnFilter"); if ($btnFilter) { $btnFilter.Add_Click($ApplyFilter) }
            $btnCloseLog = $logWin.FindName("btnCloseLog"); if ($btnCloseLog) { $btnCloseLog.Add_Click({ $logWin.Close() }.GetNewClosure()) }
            $btnExport = $logWin.FindName("btnExport")
            if ($btnExport) {
                $btnExport.Add_Click({
                    $sfd = New-Object Microsoft.Win32.SaveFileDialog; $sfd.Filter = "CSV (*.csv)|*.csv"; $sfd.FileName = "Export_Logs.csv"
                    if ($sfd.ShowDialog() -eq $true -and $lvLogs) { 
                        $lvLogs.ItemsSource | Export-Csv -Path $sfd.FileName -NoTypeInformation
                        Show-AppMessageBox -Message "Exported." -Title "Success" -ThemeColors $colors
                    }
                }.GetNewClosure())
            }
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-CenteredOnOwner -ChildWindow $logWin -OwnerWindow $Window
            $logWin.Show()
        }.GetNewClosure())
    }

    $ShowDetailsAction = {
        if ($lvData.SelectedItem) {
            $det = Get-UserDetails -Identity $lvData.SelectedItem.Name -Type $lvData.SelectedItem.Type -Config $Config
            if ($det) {
                $ftToStr = {
                    param($val)
                    if ($null -eq $val) { return "" }
                    $intVal = $val -as [Int64]
                    if ($null -eq $intVal -or $intVal -le 0 -or $intVal -ge 9223372036854770000) { return "Never" }
                    try { return [datetime]::FromFileTime($intVal).ToString("MM/dd/yyyy h:mm tt") } catch { return "" }
                }

                $panelTitle = $Window.FindName("lblDetailsPanelTitle")
                $panelSub   = $Window.FindName("lblDetailsPanelSub")
                if ($panelTitle) { $panelTitle.Text = if ($det.DisplayName) { $det.DisplayName } else { $lvData.SelectedItem.Name } }
                if ($panelSub)   { $panelSub.Text   = if ($det.DistinguishedName) { ($det.DistinguishedName -split ',(?=OU|DC)' | Select-Object -Skip 1) -join ', ' } else { "" } }

                $fieldMap = @{
                    "txtDet_DisplayName"    = $det.DisplayName
                    "txtDet_SAM"            = $det.SamAccountName
                    "txtDet_EmployeeID"     = $det.EmployeeID
                    "txtDet_EmployeeNumber" = $det.EmployeeNumber
                    "txtDet_Email"          = $det.EmailAddress
                    "txtDet_Department"     = $det.Department
                    "txtDet_Title"          = $det.Title
                    "txtDet_Phone"          = $det.OfficePhone
                    "txtDet_LastLogon"      = (& $ftToStr $det.LastLogon)
                    "txtDet_PwdLastSet"     = (& $ftToStr $det.PwdLastSet)
                    "txtDet_OU"             = $det.DistinguishedName
                }
                foreach ($name in $fieldMap.Keys) {
                    $ctrl = $Window.FindName($name)
                    if ($ctrl) { $ctrl.Text = if ($fieldMap[$name]) { "$($fieldMap[$name])" } else { "" } }
                }

                $mgrCtrl = $Window.FindName("txtDet_Manager")
                if ($mgrCtrl) {
                    $mgrCtrl.Text = if ($det.Manager) { ($det.Manager -replace '^CN=([^,]+).*', '$1') } else { "" }
                }

                $statusCtrl = $Window.FindName("txtDet_AccountStatus")
                if ($statusCtrl) {
                    $locked  = $det.LockedOut -eq $true
                    $enabled = $det.Enabled   -eq $true
                    if ($locked)        { $statusCtrl.Text = "Locked Out"; $statusCtrl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#EF4444")) }
                    elseif (-not $enabled) { $statusCtrl.Text = "Disabled";   $statusCtrl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#F59E0B")) }
                    else                { $statusCtrl.Text = "Active";     $statusCtrl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#22C55E")) }
                }

                $pwdExpiryCtrl = $Window.FindName("txtDet_PwdExpiry")
                if ($pwdExpiryCtrl) {
                    if ($det.PasswordNeverExpires -eq $true) {
                        $pwdExpiryCtrl.Text = "Never"
                        $pwdExpiryCtrl.Foreground = $Window.Resources["TextSecondary"]
                    } else {
                        $setFt = $det.PwdLastSet -as [Int64]
                        if ($setFt -and $setFt -gt 0) {
                            try {
                                $domain = Get-ADDomain -Server $Config.GeneralSettings.DomainName -ErrorAction SilentlyContinue
                                $maxAge = if ($domain) { $domain.MaxPasswordAge } else { New-TimeSpan -Days 90 }
                                $expiry = [datetime]::FromFileTime($setFt) + $maxAge
                                $daysLeft = [math]::Ceiling(($expiry - (Get-Date)).TotalDays)
                                $pwdExpiryCtrl.Text = "$($expiry.ToString('MM/dd/yyyy')) ($daysLeft days)"
                                $color = if ($daysLeft -le 7) { "#EF4444" } elseif ($daysLeft -le 14) { "#F59E0B" } else { "#22C55E" }
                                $pwdExpiryCtrl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($color))
                            } catch { $pwdExpiryCtrl.Text = "Unknown" }
                        } else { $pwdExpiryCtrl.Text = "" }
                    }
                }

                $acctExpiryCtrl = $Window.FindName("txtDet_AccountExpiry")
                if ($acctExpiryCtrl) { $acctExpiryCtrl.Text = (& $ftToStr $det.AccountExpires) }

                if ($txtDetailsContent) {
                    $displayObj = [ordered]@{}
                    $dateProps  = @("accountexpires","badpasswordtime","lastlogon","lastlogontimestamp","lockouttime","pwdlastset")
                    $visibleProps = $det.PSObject.Properties | Where-Object {
                        $_.Name -notmatch "^(PropertyNames|AddedProperties|RemovedProperties|ModifiedProperties|ClearProperties|SessionInfo)$" -and
                        $_.Name -notmatch "Certificate"
                    } | Sort-Object Name
                    foreach ($p in $visibleProps) {
                        $val = $p.Value
                        if ($null -ne $val -and $p.Name.ToLower() -in $dateProps) {
                            $valInt = $val -as [Int64]
                            if ($null -ne $valInt) {
                                if ($valInt -eq 0 -or $valInt -ge 9223372036854770000) { $val = "Never" }
                                else { try { $val = [datetime]::FromFileTime($valInt).ToString("MM/dd/yyyy h:mm:ss tt") } catch {} }
                            }
                        }
                        $displayObj[$p.Name] = $val
                    }
                    $txtDetailsContent.Text = ([PSCustomObject]$displayObj | Format-List | Out-String)
                }

                if ($overlayDetails) { $overlayDetails.Visibility = "Visible" }
            }
        }
    }.GetNewClosure()

    if ($ctxDetails) { $ctxDetails.Add_Click($ShowDetailsAction) }
    if ($lvData) { $lvData.Add_MouseDoubleClick($ShowDetailsAction) }

    $CloseDetailsPanel = { if ($overlayDetails) { $overlayDetails.Visibility = "Collapsed" } }.GetNewClosure()

    if ($btnCloseDetails) {
        $btnCloseDetails.Add_Click($CloseDetailsPanel)
    }

    $btnCloseDetailsFooter = $Window.FindName("btnCloseDetailsFooter")
    if ($btnCloseDetailsFooter) {
        $btnCloseDetailsFooter.Add_Click($CloseDetailsPanel)
    }

    if ($overlayDetails) {
        $overlayDetails.Add_MouseDown({
            param($sender, $e)
            if ($e.Source -eq $overlayDetails) {
                $overlayDetails.Visibility = "Collapsed"
            }
        }.GetNewClosure())
    }

    if ($thumbResizeDetails) {
        $thumbResizeDetails.Add_DragDelta({
            param($sender, $e)
            $newWidth = $borderDetails.Width - $e.HorizontalChange
            if ($newWidth -ge 280 -and $newWidth -le 650) { $borderDetails.Width = $newWidth }
        }.GetNewClosure())
    }

    if ($btnDashboard) { $btnDashboard.Add_Click({ Start-Process "msedge.exe" -ArgumentList "--app=""https://vm-simplify/simplifyit/custom/fileuploads/acctdashboard.html""" }.GetNewClosure()) }

    if ($btnOpenDashboard) {
        $btnOpenDashboard.Add_Click({
            $colors = Get-FluentThemeColors $State
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

            $allLogs = @(Get-AppLogFiles -Config $Config)
            $now = Get-Date

            $unlocksByDay = @{}
            for ($i = 13; $i -ge 0; $i--) {
                $day = $now.AddDays(-$i).ToString("yyyy-MM-dd")
                $unlocksByDay[$day] = 0
            }
            $unlockEvents = @($allLogs | Where-Object { $_.Event -match "Unlock" })
            foreach ($e in $unlockEvents) {
                try {
                    $day = ([datetime]$e.Timestamp).ToString("yyyy-MM-dd")
                    if ($unlocksByDay.ContainsKey($day)) { $unlocksByDay[$day]++ }
                } catch {}
            }

            $cutoff = $now.AddDays(-30)
            $recentLockouts = @($allLogs | Where-Object {
                if ($_.Event -notmatch "Lockout Detected") { return $false }
                $parsed = [datetime]::MinValue
                if (-not [datetime]::TryParse($_.Timestamp, [ref]$parsed)) { return $false }
                return $parsed -gt $cutoff
            })
            $topLocked = $recentLockouts | Group-Object Username | Sort-Object Count -Descending | Select-Object -First 5

            $monthStart = [datetime]::new($now.Year, $now.Month, 1)
            $thisMonth = @($allLogs | Where-Object {
                $parsed = [datetime]::MinValue
                if (-not [datetime]::TryParse($_.Timestamp, [ref]$parsed)) { return $false }
                return $parsed -ge $monthStart
            })
            $totalUnlocksMonth   = @($thisMonth | Where-Object { $_.Event -match "Unlock Account" }).Count
            $totalResetsMonth    = @($thisMonth | Where-Object { $_.Event -match "Password Reset" }).Count
            $totalPrinterMonth   = @($thisMonth | Where-Object { $_.Event -match "Printer" }).Count
            $totalTicketsMonth   = @($thisMonth | Where-Object { $_.Event -match "Freshservice" }).Count

            [System.Windows.Input.Mouse]::OverrideCursor = $null

            $dashXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Management Dashboard" Width="720" SizeToContent="Height"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
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
    <Border Background="$($colors.Bg)" CornerRadius="10" BorderBrush="$($colors.BtnBorder)" BorderThickness="1" Margin="15">
        <Border.Effect><DropShadowEffect BlurRadius="24" ShadowDepth="6" Opacity="0.28" Color="Black"/></Border.Effect>
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="22,20,22,0" Cursor="Hand">
                <StackPanel>
                    <TextBlock Text="Management Dashboard" FontSize="18" FontWeight="SemiBold" Foreground="$($colors.Fg)"/>
                    <TextBlock Text="Activity summary -- $($now.ToString('MMMM yyyy'))" FontSize="12" Foreground="$($colors.SecFg)" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>

            <StackPanel Grid.Row="1" Margin="22,16,22,16">

                <!-- KPI row -->
                <Grid Margin="0,0,0,14">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" Padding="14,12" Background="$($colors.BtnBg)" CornerRadius="8" BorderBrush="$($colors.GridBorder)" BorderThickness="1">
                        <StackPanel HorizontalAlignment="Center">
                            <TextBlock Text="$totalUnlocksMonth" FontSize="32" FontWeight="Bold" Foreground="$($colors.PrimaryBg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="Unlocks" FontSize="12" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="This Month" FontSize="11" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Column="2" Padding="14,12" Background="$($colors.BtnBg)" CornerRadius="8" BorderBrush="$($colors.GridBorder)" BorderThickness="1">
                        <StackPanel HorizontalAlignment="Center">
                            <TextBlock Text="$totalResetsMonth" FontSize="32" FontWeight="Bold" Foreground="$($colors.PrimaryBg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="Resets" FontSize="12" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="This Month" FontSize="11" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Column="4" Padding="14,12" Background="$($colors.BtnBg)" CornerRadius="8" BorderBrush="$($colors.GridBorder)" BorderThickness="1">
                        <StackPanel HorizontalAlignment="Center">
                            <TextBlock Text="$totalPrinterMonth" FontSize="32" FontWeight="Bold" Foreground="$($colors.PrimaryBg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="Printer Ops" FontSize="12" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="This Month" FontSize="11" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Column="6" Padding="14,12" Background="$($colors.BtnBg)" CornerRadius="8" BorderBrush="$($colors.GridBorder)" BorderThickness="1">
                        <StackPanel HorizontalAlignment="Center">
                            <TextBlock Text="$totalTicketsMonth" FontSize="32" FontWeight="Bold" Foreground="$($colors.PrimaryBg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="FS Tickets" FontSize="12" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                            <TextBlock Text="This Month" FontSize="11" Foreground="$($colors.SecFg)" HorizontalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- 14-day unlock bar chart -->
                <Border Padding="14,12" Background="$($colors.BtnBg)" CornerRadius="8" BorderBrush="$($colors.GridBorder)" BorderThickness="1" Margin="0,0,0,14">
                    <StackPanel>
                        <TextBlock Text="Unlocks -- Last 14 Days" FontSize="13" FontWeight="SemiBold" Foreground="$($colors.SecFg)" Margin="0,0,0,10"/>
                        <Grid x:Name="chartGrid" Height="80">
                        </Grid>
                        <Grid x:Name="chartLabels" Height="18" Margin="0,2,0,0">
                        </Grid>
                    </StackPanel>
                </Border>

                <!-- Top locked users -->
                <Border Padding="14,12" Background="$($colors.BtnBg)" CornerRadius="8" BorderBrush="$($colors.GridBorder)" BorderThickness="1">
                    <StackPanel>
                        <TextBlock Text="Most Locked Accounts -- Last 30 Days" FontSize="13" FontWeight="SemiBold" Foreground="$($colors.SecFg)" Margin="0,0,0,10"/>
                        <StackPanel x:Name="icTopLocked"/>
                        <TextBlock x:Name="lblNoLockouts" Text="No lockout events in the last 30 days." FontSize="12" Foreground="$($colors.SecFg)" Visibility="Collapsed"/>
                    </StackPanel>
                </Border>
            </StackPanel>

            <Border Grid.Row="2" Background="$($colors.BtnBg)" CornerRadius="0,0,10,10" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($colors.BtnBorder)">
                <Button x:Name="btnCloseDash" Content="Close" HorizontalAlignment="Right" Width="80" Height="28" Background="$($colors.BtnBg)" Foreground="$($colors.Fg)" BorderBrush="$($colors.BtnBorder)" BorderThickness="1"/>
            </Border>
        </Grid>
    </Border>
</Window>
"@
            $dReader  = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dashXaml))
            $dashWin  = [System.Windows.Markup.XamlReader]::Load($dReader)
            $dashWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $dashWin.DragMove() })
            $dashWin.FindName("btnCloseDash").Add_Click({ $dashWin.Close() }.GetNewClosure())

            $dashWin.Add_Loaded({
                $chartGrid   = $dashWin.FindName("chartGrid")
                $chartLabels = $dashWin.FindName("chartLabels")

                if ($chartGrid) {
                    $days   = @($unlocksByDay.Keys | Sort-Object)
                    $maxVal = ($unlocksByDay.Values | Measure-Object -Maximum).Maximum
                    if ($maxVal -lt 1) { $maxVal = 1 }

                    $accentR = 59; $accentG = 130; $accentB = 246
                    try {
                        $parsed = [System.Windows.Media.ColorConverter]::ConvertFromString($colors.PrimaryBg)
                        $accentR = $parsed.R; $accentG = $parsed.G; $accentB = $parsed.B
                    } catch {}

                    for ($i = 0; $i -lt $days.Count; $i++) {
                        $cd1 = New-Object System.Windows.Controls.ColumnDefinition
                        $cd1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                        $chartGrid.ColumnDefinitions.Add($cd1)

                        $cd2 = New-Object System.Windows.Controls.ColumnDefinition
                        $cd2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                        $chartLabels.ColumnDefinitions.Add($cd2)

                        $d   = $days[$i]
                        $val = $unlocksByDay[$d]
                        $pct = if ($maxVal -gt 0) { $val / $maxVal } else { 0.0 }
                        $barH = if ($val -eq 0) { 2 } else { [math]::Max(4, [int](76 * $pct)) }

                        $t  = $pct
                        $r  = [int](180 + ($accentR - 180) * $t)
                        $g  = [int](180 + ($accentG - 180) * $t)
                        $b  = [int](180 + ($accentB - 180) * $t)
                        $barColor = "#{0:X2}{1:X2}{2:X2}" -f [math]::Max(0,[math]::Min(255,$r)), `
                                                               [math]::Max(0,[math]::Min(255,$g)), `
                                                               [math]::Max(0,[math]::Min(255,$b))

                        $bar = New-Object System.Windows.Shapes.Rectangle
                        $bar.Fill             = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($barColor))
                        $bar.RadiusX          = 3
                        $bar.RadiusY          = 3
                        $bar.Margin           = [System.Windows.Thickness]::new(3, 0, 3, 0)
                        $bar.VerticalAlignment   = [System.Windows.VerticalAlignment]::Bottom
                        $bar.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
                        $bar.Height = $barH
                        $suffix = if ($val -ne 1) { "s" } else { "" }
                        $bar.ToolTip = "$val unlock$suffix on $d"
                        [System.Windows.Controls.Grid]::SetColumn($bar, $i)
                        $chartGrid.Children.Add($bar) | Out-Null

                        if ($i % 2 -eq 0) {
                            $lbl = New-Object System.Windows.Controls.TextBlock
                            $lbl.Text                = ([datetime]$d).ToString("M/d")
                            $lbl.FontSize            = 9
                            $lbl.Foreground          = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.SecFg))
                            $lbl.HorizontalAlignment = "Center"
                            $lbl.VerticalAlignment   = "Center"
                            [System.Windows.Controls.Grid]::SetColumn($lbl, $i)
                            $chartLabels.Children.Add($lbl) | Out-Null
                        }
                    }
                }

                $icTopLocked   = $dashWin.FindName("icTopLocked")
                $lblNoLockouts = $dashWin.FindName("lblNoLockouts")

                if ($topLocked -and @($topLocked).Count -gt 0) {
                    $topList  = @($topLocked)
                    $maxCount = ($topList | Measure-Object Count -Maximum).Maximum
                    if ($maxCount -lt 1) { $maxCount = 1 }

                    $rankColors = @("#EF4444", "#F59E0B", $colors.PrimaryBg, $colors.PrimaryBg, $colors.PrimaryBg)

                    for ($ri = 0; $ri -lt $topList.Count; $ri++) {
                        $entry    = $topList[$ri]
                        $barColor = if ($ri -lt $rankColors.Count) { $rankColors[$ri] } else { $colors.PrimaryBg }

                        $row = New-Object System.Windows.Controls.Grid
                        $row.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

                        $c0 = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = [System.Windows.GridLength]::new(160)
                        $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                        $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = [System.Windows.GridLength]::new(36)
                        $row.ColumnDefinitions.Add($c0)
                        $row.ColumnDefinitions.Add($c1)
                        $row.ColumnDefinitions.Add($c2)

                        $nameLbl = New-Object System.Windows.Controls.TextBlock
                        $nameLbl.Text               = $entry.Name
                        $nameLbl.Foreground         = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.Fg))
                        $nameLbl.FontSize            = 13
                        $nameLbl.VerticalAlignment   = "Center"
                        $nameLbl.TextTrimming        = "CharacterEllipsis"
                        [System.Windows.Controls.Grid]::SetColumn($nameLbl, 0)
                        $row.Children.Add($nameLbl) | Out-Null

                        $trackGrid = New-Object System.Windows.Controls.Grid
                        $trackGrid.VerticalAlignment = "Center"

                        $track = New-Object System.Windows.Shapes.Rectangle
                        $track.Height  = 10
                        $track.RadiusX = 3; $track.RadiusY = 3
                        $track.Fill    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.GridBorder))
                        $track.HorizontalAlignment = "Stretch"
                        $trackGrid.Children.Add($track) | Out-Null

                        $bar = New-Object System.Windows.Shapes.Rectangle
                        $bar.Height  = 10
                        $bar.RadiusX = 3; $bar.RadiusY = 3
                        $bar.Fill    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($barColor))
                        $bar.HorizontalAlignment = "Left"
                        $bar.Tag = $entry.Count / $maxCount
                        $trackGrid.Children.Add($bar) | Out-Null

                        $barRef   = $bar
                        $trackRef = $trackGrid
                        $trackGrid.Add_SizeChanged({
                            param($s, $e)
                            $barRef.Width = [math]::Max(4, $trackRef.ActualWidth * $barRef.Tag)
                        }.GetNewClosure())

                        [System.Windows.Controls.Grid]::SetColumn($trackGrid, 1)
                        $row.Children.Add($trackGrid) | Out-Null

                        $countLbl = New-Object System.Windows.Controls.TextBlock
                        $countLbl.Text               = $entry.Count
                        $countLbl.Foreground         = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($barColor))
                        $countLbl.FontSize            = 12
                        $countLbl.FontWeight          = "SemiBold"
                        $countLbl.HorizontalAlignment = "Right"
                        $countLbl.VerticalAlignment   = "Center"
                        [System.Windows.Controls.Grid]::SetColumn($countLbl, 2)
                        $row.Children.Add($countLbl) | Out-Null

                        $icTopLocked.Children.Add($row) | Out-Null
                    }
                } else {
                    if ($icTopLocked)   { $icTopLocked.Visibility   = "Collapsed" }
                    if ($lblNoLockouts) { $lblNoLockouts.Visibility = "Visible" }
                }
            }.GetNewClosure())

            Show-CenteredOnOwner -ChildWindow $dashWin -OwnerWindow $Window
            $dashWin.ShowDialog() | Out-Null
        }.GetNewClosure())
    }
    if ($btnHelp) { $btnHelp.Add_Click({ 
        $helpPath = Join-Path $AppRoot "HDCompanionUserGuide.html"
        if (Test-Path $helpPath) { Start-Process "msedge.exe" -ArgumentList "--app=""$helpPath""" } else { Show-AppMessageBox -Message "Help document not found." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) | Out-Null }
    }.GetNewClosure()) }
    if ($btnTechDocs) { $btnTechDocs.Add_Click({ 
        $techPath = Join-Path $AppRoot "HDCompanionTechDoc.html"
        if (Test-Path $techPath) { Start-Process "msedge.exe" -ArgumentList "--app=""$techPath""" } else { Show-AppMessageBox -Message "Tech document not found." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) | Out-Null }
    }.GetNewClosure()) }
    if ($btnModifyConfig) {
        $btnModifyConfig.Add_Click({
            $editorPath = Join-Path $AppRoot "ConfigEditor.html"
            if (Test-Path $editorPath) {
                $rawHtml = Get-Content -Path $editorPath -Raw
                $jsonStr = $Config | ConvertTo-Json -Depth 10 -Compress
                $rawHtml = $rawHtml -replace 'window\.INJECTED_CONFIG\s*=\s*null;', "window.INJECTED_CONFIG = $jsonStr;"
                $tempPath = Join-Path $env:TEMP "HDCompanion_ConfigEditor.html"
                Set-Content -Path $tempPath -Value $rawHtml -Force
                Start-Process "msedge.exe" -ArgumentList "--app=""$tempPath"""
                $res = Show-AppMessageBox -Message "Configuration Editor launched in Edge.`n`nPlease save your changes to 'AcctMonitorCfg.json' in the script directory.`n`nClick OK to reload settings now." -Title "Edit Configuration" -ButtonType "OKCancel" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                if ($res -eq "OK") {
                    $newConfig = Get-AppConfig
                    $Config.PSObject.Properties | ForEach-Object { $Config.($_.Name) = $newConfig.($_.Name) }
                    & $ApplyTheme -TargetTheme $State.CurrentTheme
                    if ($Config.ControlProperties) {
                        $lblTitle = $Window.FindName("lblTitle"); if ($lblTitle) { $lblTitle.Text = $Config.ControlProperties.TitleLabel.Text }
                        $lblSubtitle = $Window.FindName("lblSubtitle"); if ($lblSubtitle) { $lblSubtitle.Text = $Config.ControlProperties.SubtitleLabel.Text }
                        if ($btnRefresh) { $btnRefresh.Content = $Config.ControlProperties.RefreshButton.Text }
                        $btnUnlock = $Window.FindName("btnUnlock"); if ($btnUnlock) { $btnUnlock.Content = $Config.ControlProperties.UnlockButton.Text }
                        $btnUnlockAll = $Window.FindName("btnUnlockAll"); if ($btnUnlockAll) { $btnUnlockAll.Content = $Config.ControlProperties.UnlockAllButton.Text }
                        if ($btnSearch) { $btnSearch.Content = $Config.ControlProperties.SearchButton.Text }
                        if ($btnViewLog) { $btnViewLog.Content = $Config.ControlProperties.ViewLogButton.Text }
                    }
                    if ($chkEnableEmail) { $chkEnableEmail.IsChecked = $Config.EmailSettings.EnableEmailNotifications }
                    $cfgPath = if ($newConfig.LoadedConfigPath) { $newConfig.LoadedConfigPath } else { "Unknown" }
                    Add-AppLog -Event "System" -Username "System" -Details "Configuration reloaded from: $cfgPath" -Config $Config -State $State -Status "Info" -Color "Blue"
                }
            } else { Show-AppMessageBox -Message "ConfigEditor.html not found." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
        }.GetNewClosure())
    }

    $State.Timer = [System.Windows.Threading.DispatcherTimer]::new()
    $State.Timer.Add_Tick({ 
        $overlayReset = $Window.FindName("overlayReset")
        if ($overlayReset -and $overlayReset.Visibility -eq "Visible" -or ($overlayDetails -and $overlayDetails.Visibility -eq "Visible")) { return }
        $diff = $State.RefreshTargetTime - (Get-Date)
        if ($diff.TotalSeconds -le 0) { & $RefreshAction } 
        elseif ($btnRefresh) { $btnRefresh.Content = "Refresh List ({0}s)" -f [math]::Ceiling($diff.TotalSeconds) }
    }.GetNewClosure())

    $Window.Add_Closing({
        if ($State.Timer)           { $State.Timer.Stop() }
        if ($State.AutoUnlockTimer) { $State.AutoUnlockTimer.Stop() }
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }.GetNewClosure())
    
    if ($cbAutoRefresh) {
        $cbAutoRefresh.Add_SelectionChanged({
            $sec = 180; switch ($cbAutoRefresh.SelectedItem) { "30 seconds"{$sec=30}; "1 minute"{$sec=60}; "2 minutes"{$sec=120}; "5 minutes"{$sec=300}; "10 minutes"{$sec=600}; "15 minutes"{$sec=900}; "30 minutes"{$sec=1800} }
            $State.RefreshIntervalSeconds = $sec; $State.RefreshTargetTime = (Get-Date).AddSeconds($sec)
        }.GetNewClosure())
    }

    # --- THEME CUSTOMIZER LOGIC ---
    $btnPreferences = $Window.FindName("btnPreferences")
    if ($btnPreferences) {
        $btnPreferences.Add_Click({
            $ApplyThemeRef = $ApplyTheme
            $WindowRef = $Window
            $ConfigRef = $Config
            $StateRef = $State

            $xamlPath = Join-Path (Split-Path $PSScriptRoot -Parent) "UI\Dialogs\ThemeCustomizer.xaml"
            if (-not (Test-Path $xamlPath)) {
                Show-AppMessageBox -Message "ThemeCustomizer.xaml not found." -Title "Error" -IconType "Error" -OwnerWindow $WindowRef -ThemeColors (Get-FluentThemeColors $StateRef)
                return
            }

            $themeWin = Load-XamlWindow -XamlPath $xamlPath -ThemeColors (Get-FluentThemeColors $StateRef)
            $themeWin.Owner = $WindowRef
            
            $titleBar = $themeWin.FindName("TitleBar")
            if ($titleBar) { $titleBar.Add_MouseLeftButtonDown({ $themeWin.DragMove() }) }

            $cbTheme = $themeWin.FindName("cbTheme")
            $cbAccent = $themeWin.FindName("cbAccent")
            $txtCustomHex = $themeWin.FindName("txtCustomHex")
            $cbDensity = $themeWin.FindName("cbDensity")
            $cbFontSize = $themeWin.FindName("cbFontSize")
            $chkGlassEffect = $themeWin.FindName("chkGlassEffect")
            $btnSavePrefs = $themeWin.FindName("btnSavePrefs")
            $btnCancelPrefs = $themeWin.FindName("btnCancelPrefs")
            
            $txtAppImage = $themeWin.FindName("txtAppImage")
            $btnBrowseApp = $themeWin.FindName("btnBrowseApp")
            $btnClearApp = $themeWin.FindName("btnClearApp")

            # --- SEARCH HISTORY ELEMENTS ---
            $txtSearchHistory = $themeWin.FindName("txtSearchHistory")
            $btnClearHistory = $themeWin.FindName("btnClearHistory")

            $fieldFgBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Colors]::Black)
            foreach ($field in @($cbTheme, $cbAccent, $txtCustomHex, $cbDensity, $cbFontSize, $txtAppImage, $txtSearchHistory)) {
                if ($null -ne $field) { $field.Foreground = $fieldFgBrush }
            }

            # Map existing history into the multi-line textbox
            if ($txtSearchHistory -and $StateRef.SearchHistory) {
                $txtSearchHistory.Text = $StateRef.SearchHistory -join "`r`n"
            }

            if ($btnClearHistory) {
                $btnClearHistory.Add_Click({
                    if ($txtSearchHistory) { $txtSearchHistory.Text = "" }
                }.GetNewClosure())
            }

            $cbAccent.Add_SelectionChanged({
                if ($cbAccent.SelectedIndex -eq 4) { $txtCustomHex.Visibility = "Visible" } else { $txtCustomHex.Visibility = "Collapsed" }
            }.GetNewClosure())

            if ($txtCustomHex) {
                $txtCustomHex.ToolTip = "Double-click to open the Color Picker, or type a Hex code manually."
                $txtCustomHex.Add_PreviewMouseDoubleClick({
                    param($sender, $e)
                    try {
                        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
                        $colorDialog = New-Object System.Windows.Forms.ColorDialog
                        $colorDialog.FullOpen = $true
                        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $txtCustomHex.Text = "#{0:X2}{1:X2}{2:X2}" -f $colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B
                        }
                        $e.Handled = $true
                    } catch {}
                }.GetNewClosure())
            }
            
            if ($btnBrowseApp -and $txtAppImage) {
                $btnBrowseApp.Add_Click({
                    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
                    $ofd = New-Object System.Windows.Forms.OpenFileDialog
                    $ofd.Title = "Select Application Background Image"
                    $ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif"
                    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtAppImage.Text = $ofd.FileName }
                }.GetNewClosure())
            }
            if ($btnClearApp -and $txtAppImage) {
                $btnClearApp.Add_Click({ $txtAppImage.Text = "" }.GetNewClosure())
            }

            if ($StateRef.CurrentTheme -eq "Dark") { $cbTheme.SelectedIndex = 1 } else { $cbTheme.SelectedIndex = 0 }
            
            if ($ConfigRef.UserPreferences) {
                if ($ConfigRef.UserPreferences.Density -eq "Compact") { $cbDensity.SelectedIndex = 1 } else { $cbDensity.SelectedIndex = 0 }
                if ($ConfigRef.UserPreferences.FontSize -eq "Small") { $cbFontSize.SelectedIndex = 0 }
                elseif ($ConfigRef.UserPreferences.FontSize -eq "Large") { $cbFontSize.SelectedIndex = 2 }
                else { $cbFontSize.SelectedIndex = 1 }
                $chkGlassEffect.IsChecked = ($ConfigRef.UserPreferences.GlassEffect -eq $true)
                
                if ($txtAppImage -and $ConfigRef.UserPreferences.AppBackgroundImage) { $txtAppImage.Text = $ConfigRef.UserPreferences.AppBackgroundImage }
            }
            
            $primaryVal = $ConfigRef.LightModeColors.Primary
            if ($primaryVal -is [string]) {
                $cbAccent.SelectedIndex = 4
                $txtCustomHex.Text = $primaryVal
                $txtCustomHex.Visibility = "Visible"
            } else {
                $primaryStr = $primaryVal -join ","
                switch ($primaryStr) {
                    "37,99,235"  { $cbAccent.SelectedIndex = 0 } 
                    "163,74,40"  { $cbAccent.SelectedIndex = 1 } 
                    "5,150,105"  { $cbAccent.SelectedIndex = 2 } 
                    "147,51,234" { $cbAccent.SelectedIndex = 3 } 
                    default      { 
                        $cbAccent.SelectedIndex = 4 
                        $txtCustomHex.Visibility = "Visible"
                        try { $txtCustomHex.Text = "#{0:X2}{1:X2}{2:X2}" -f $primaryVal[0], $primaryVal[1], $primaryVal[2] } catch {}
                    }
                }
            }

            $btnCancelPrefs.Add_Click({ $themeWin.Close() })

            $btnSavePrefs.Add_Click({
                $newTheme = if ($cbTheme.SelectedIndex -eq 1) { "Dark" } else { "Light" }
                $ConfigRef.GeneralSettings.DefaultTheme = $newTheme
                $StateRef.CurrentTheme = $newTheme

                $pLight = $null; $pDark = $null
                switch ($cbAccent.SelectedIndex) {
                    0 { $pLight = @(37, 99, 235); $pDark = @(59, 130, 246) }
                    1 { $pLight = @(163, 74, 40); $pDark = @(216, 140, 108) }
                    2 { $pLight = @(5, 150, 105); $pDark = @(16, 185, 129) }
                    3 { $pLight = @(147, 51, 234); $pDark = @(168, 85, 247) }
                    4 { 
                        $hex = $txtCustomHex.Text.Trim()
                        if (-not $hex.StartsWith("#")) { $hex = "#$hex" }
                        $pLight = $hex; $pDark = $hex 
                    }
                }

                $ConfigRef.LightModeColors.Primary = $pLight
                $ConfigRef.DarkModeColors.Primary = $pDark

                $density = if ($cbDensity.SelectedIndex -eq 1) { "Compact" } else { "Comfortable" }
                $fontSize = switch ($cbFontSize.SelectedIndex) { 0 {"Small"} 2 {"Large"} default {"Default"} }
                $glass = ($chkGlassEffect.IsChecked -eq $true)
                $appImg = if ($txtAppImage) { $txtAppImage.Text.Trim() } else { $ConfigRef.UserPreferences.AppBackgroundImage }

                if (-not $ConfigRef.UserPreferences) { $ConfigRef | Add-Member -MemberType NoteProperty -Name "UserPreferences" -Value @{} -Force }
                $ConfigRef.UserPreferences.Density = $density
                $ConfigRef.UserPreferences.FontSize = $fontSize
                $ConfigRef.UserPreferences.GlassEffect = $glass
                $ConfigRef.UserPreferences.AppBackgroundImage = $appImg

                # --- Save the modified Search History ---
                if ($txtSearchHistory) {
                    $parsedHistory = $txtSearchHistory.Text -split "\r?\n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Select-Object -First 20
                    $StateRef.SearchHistory = @($parsedHistory)

                    $mainSearchBox = $WindowRef.FindName("txtSearch")
                    if ($mainSearchBox) {
                        $mainSearchBox.ItemsSource = $StateRef.SearchHistory
                    }
                }

                $prefsDir = Join-Path $env:LOCALAPPDATA "PelicanCU\HDCompanion"
                if (-not (Test-Path $prefsDir)) { New-Item -ItemType Directory -Path $prefsDir -Force | Out-Null }
                
                $prefsObj = @{
                    DefaultTheme = $newTheme
                    LightModeColors = @{ Primary = $pLight }
                    DarkModeColors = @{ Primary = $pDark }
                    Density = $density
                    FontSize = $fontSize
                    GlassEffect = $glass
                    SearchHistory = $StateRef.SearchHistory
                    AppBackgroundImage = $appImg
                }
                
                $prefsObj | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $prefsDir "userprefs.json") -Encoding UTF8

                & $ApplyThemeRef -TargetTheme $newTheme
                $themeWin.Close()
            }.GetNewClosure())

            Show-CenteredOnOwner -ChildWindow $themeWin -OwnerWindow $Window
            $themeWin.ShowDialog() | Out-Null
        }.GetNewClosure())
    }
}

Export-ModuleMember -Function *