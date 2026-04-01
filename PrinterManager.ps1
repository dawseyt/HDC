# ============================================================================
# PrinterManager.ps1 - Standalone Printer Management Thread
# ============================================================================
param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME,

    # Accept the runtime theme from the parent process so the window matches
    # the user's active theme, even if they toggled it after startup.
    # Falls back to the DefaultTheme from config if not supplied.
    [Parameter(Mandatory=$false)]
    [string]$Theme = ""
)

# 1. Environment Setup
try {
    $global:PSScriptRoot = $PSScriptRoot
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
} catch { exit }

# Ensure 64-bit execution (Get-PrinterDriver often fails silently in 32-bit processes on 64-bit OS)
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $psExe = "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe"
    $argList = @("-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"", "-ComputerName", "`"$ComputerName`"", "-Theme", "`"$Theme`"")
    Start-Process $psExe -ArgumentList $argList
    exit
}

# 2. Import Required Backend Modules
Import-Module "$PSScriptRoot\Modules\CoreLogic.psm1" -Force -DisableNameChecking
Import-Module "$PSScriptRoot\Modules\RemoteManagement.psm1" -Force -DisableNameChecking

# 3. Setup Config and Theme Colors
$Config = Get-AppConfig

# Expose config as $global:Config so Get-FluentThemeColors can read the
# user's accent color preference -- it checks $global:Config internally.
$global:Config = $Config

# Determine theme: prefer the runtime value passed from the parent process
# (reflects any mid-session toggle), fall back to the config file default.
$activeTheme = if ($Theme -ne "") { $Theme } else { $Config.GeneralSettings.DefaultTheme }
$isDark = ($activeTheme -eq "Dark")

# Build the color map using the same function as the main application,
# so accent color and all palette values stay in sync.
$pmState = @{ CurrentTheme = $activeTheme }
$c = Get-FluentThemeColors -State $pmState

# ---------------------------------------------------------------------------
# Center-OnPrnWin
# Centers a child window on the same monitor as $prnWin. Because PrinterManager
# is a separate process, it cannot call Show-CenteredOnOwner from UI_Core.
# ---------------------------------------------------------------------------
function Center-OnPrnWin {
    param($ChildWindow, $OwnerWindow)
    try {
        $centrePt = [System.Drawing.Point]::new(
            [int]($OwnerWindow.Left + $OwnerWindow.ActualWidth  / 2),
            [int]($OwnerWindow.Top  + $OwnerWindow.ActualHeight / 2)
        )
        $screen = [System.Windows.Forms.Screen]::FromPoint($centrePt)
        $wa     = $screen.WorkingArea
        $ChildWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
        $childW = if ($ChildWindow.Width  -gt 0) { $ChildWindow.Width  } else { 450 }
        $childH = if ($ChildWindow.Height -gt 0) { $ChildWindow.Height } else { 300 }
        $ChildWindow.Left = $wa.Left + ($wa.Width  - $childW) / 2
        $ChildWindow.Top  = $wa.Top  + ($wa.Height - $childH) / 2
    } catch {
        $ChildWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
        if ($OwnerWindow -and $OwnerWindow.IsVisible) { $ChildWindow.Owner = $OwnerWindow }
    }
}

# 4. Universal Message Box (Embedded XAML)
function Show-LocalMessageBox {
    param([string]$Message, [string]$Title = "Information", [string]$ButtonType = "OK", [string]$IconType = "Information", $OwnerWindow)
    
    $iconData = ""
    $iconColor = $c.Fg
    switch ($IconType) {
        "Error" { $iconData = "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"; $iconColor = "#D13438" }
        "Warning" { $iconData = "M12 2L1 21h22L12 2zm0 3.83L19.53 19H4.47L12 5.83zM11 10h2v5h-2v-5zm0 6h2v2h-2v-2z"; $iconColor = "#FFB900" }
        "Question" { $iconData = "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17h-2v-2h2v2zm2.07-7.75l-.9.92C13.45 12.9 13 13.5 13 15h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26c.37-.36.59-.86.59-1.41 0-1.1-.9-2-2-2s-2 .9-2 2H8c0-2.21 1.79-4 4-4s4 1.79 4 4c0 .88-.36 1.68-.93 2.25z"; $iconColor = "#0078D4" }
        Default { $iconData = "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h-2v6h2zm0-8h-2V7h2v2z"; $iconColor = "#0078D4" }
    }

    $btnOkVis = if ($ButtonType -in @("OK", "OKCancel")) { "Visible" } else { "Collapsed" }
    $btnCancelVis = if ($ButtonType -eq "OKCancel") { "Visible" } else { "Collapsed" }
    $btnYesVis = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }
    $btnNoVis = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }

    $msgXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="$Title" Width="450" SizeToContent="Height" MinHeight="180" 
            WindowStartupLocation="CenterOwner" ResizeMode="NoResize" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
            FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
        <Window.Resources>
            <Style TargetType="Button">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Window.Resources>
        <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3" Color="Black"/></Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid x:Name="TitleBar" Grid.Row="0" Margin="20,20,20,0" Background="Transparent" Cursor="Hand">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <Path Grid.Column="0" Data="$iconData" Fill="$iconColor" Width="20" Height="20" Stretch="Uniform" VerticalAlignment="Center"/>
                    <TextBlock Grid.Column="1" Text="$Title" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)" VerticalAlignment="Center" Margin="12,0,0,0"/>
                </Grid>
                <TextBox x:Name="txtMessageBody" Grid.Row="1" Margin="52,12,20,20" IsReadOnly="True" Background="Transparent" BorderThickness="0" TextWrapping="Wrap" Foreground="$($c.SecFg)" FontSize="13" VerticalScrollBarVisibility="Auto" MaxHeight="300"/>
                <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="btnYes" Content="Yes" Width="80" Height="28" Margin="0,0,8,0" Visibility="$btnYesVis" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0"/>
                        <Button x:Name="btnNo" Content="No" Width="80" Height="28" Margin="0,0,0,0" Visibility="$btnNoVis" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                        <Button x:Name="btnOk" Content="OK" Width="80" Height="28" Margin="0,0,8,0" Visibility="$btnOkVis" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0"/>
                        <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" Margin="0,0,0,0" Visibility="$btnCancelVis" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
    </Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($msgXaml))
    $msgWin = [System.Windows.Markup.XamlReader]::Load($reader)
    
    $msgWin.FindName("txtMessageBody").Text = $Message
    $msgWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $msgWin.DragMove() })
    
    $script:msgRes = "Cancel"
    $msgWin.FindName("btnYes").Add_Click({ $script:msgRes = "Yes"; $msgWin.Close() })
    $msgWin.FindName("btnNo").Add_Click({ $script:msgRes = "No"; $msgWin.Close() })
    $msgWin.FindName("btnOk").Add_Click({ $script:msgRes = "OK"; $msgWin.Close() })
    $msgWin.FindName("btnCancel").Add_Click({ $script:msgRes = "Cancel"; $msgWin.Close() })
    
    if ($OwnerWindow -and $OwnerWindow.IsVisible) { $msgWin.Owner = $OwnerWindow; $msgWin.WindowStartupLocation = "CenterOwner" }
    else { $msgWin.WindowStartupLocation = "CenterScreen" }
    
    $msgWin.ShowDialog() | Out-Null
    return $script:msgRes
}

# 4.5 Universal Inline Progress Function
function Invoke-WithProgress {
    param([string]$Title, $OwnerWindow, [scriptblock]$Action)
    
    $progXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="$Title" Width="400" SizeToContent="Height" WindowStartupLocation="CenterOwner" 
            WindowStyle="None" AllowsTransparency="True" Background="Transparent" Topmost="True"
            FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
        <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
            <StackPanel Margin="20">
                <TextBlock Text="$Title" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)" Margin="0,0,0,10"/>
                <ProgressBar x:Name="pbStatus" Minimum="0" Maximum="100" Value="0" Height="6" Margin="0,0,0,10" Foreground="$($c.PrimaryBg)" Background="$($c.BtnBg)" BorderThickness="0"/>
                <TextBlock x:Name="lblStatus" Text="Initializing..." FontSize="12" Foreground="$($c.SecFg)" TextWrapping="Wrap"/>
            </StackPanel>
        </Border>
    </Window>
"@
    $pReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($progXaml))
    $progWin = [System.Windows.Markup.XamlReader]::Load($pReader)
    if ($OwnerWindow -and $OwnerWindow.IsVisible) { $progWin.Owner = $OwnerWindow }
    else { $progWin.WindowStartupLocation = "CenterScreen" }
    
    $lblStatus = $progWin.FindName("lblStatus")
    $pbStatus = $progWin.FindName("pbStatus")
    $progWin.Show()
    
    # Delegate used to update status text continuously while avoiding UI hangs
    $UpdateText = {
        param([string]$Message, [int]$Progress = -1)
        $lblStatus.Text = $Message
        if ($Progress -ge 0 -and $Progress -le 100) {
            $pbStatus.Value = $Progress
        }
        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    }
    
    try {
        return & $Action $UpdateText
    } finally {
        $progWin.Close()
    }
}

# 5. Load the Unified Printer Manager XAML (Embedded)
$pmXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="winPrinterManager" Title="Printer Management" Height="640" Width="980"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" WindowStyle="SingleBorderWindow"
        Background="$($c.Bg)" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
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
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="$($c.Bg)"/>
            <Setter Property="Foreground" Value="$($c.SecFg)"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="$($c.GridBorder)"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header bar -->
        <Border Grid.Row="0" Background="$($c.Bg)" Padding="16" BorderBrush="$($c.GridBorder)" BorderThickness="0,0,0,1">
            <Grid>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock x:Name="lblHeaderTitle" Text="Printer Management: $ComputerName" FontWeight="SemiBold" FontSize="18" Foreground="$($c.Fg)"/>
                    <TextBlock x:Name="lblDescription" Text="Asset: Querying..." FontSize="13" Foreground="$($c.SecFg)" Margin="0,2,0,0"/>
                    <TextBlock x:Name="lblStatus" Text="Ready." FontSize="12" Foreground="$($c.SecFg)" Margin="0,4,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnRestartSpooler" Content="Restart Spooler"    Margin="0,0,10,0" Width="120" Height="30" Background="$($c.BtnBg)"    Foreground="$($c.Fg)"      BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    <Button x:Name="btnInstallPrinter" Content="Install New Printer"               Width="140" Height="30" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Tab area -->
        <TabControl Grid.Row="1" Margin="8,8,8,0" Background="Transparent" BorderThickness="0">

            <!-- Tab 0: Printers -->
            <TabItem Header="Printers" FontSize="13">
                <Border Background="$($c.Bg)" CornerRadius="0,6,6,6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                    <ListView x:Name="lvPrinters" BorderThickness="0" Margin="4" Background="Transparent" Foreground="$($c.Fg)" AlternationCount="2">
                        <ListView.ItemContainerStyle>
                            <Style TargetType="ListViewItem">
                                <Setter Property="Height" Value="28"/>
                                <Setter Property="Background" Value="Transparent"/>
                                <Setter Property="Foreground" Value="$($c.Fg)"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="ListViewItem">
                                            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="0" CornerRadius="4" Padding="4,0" Margin="0,1">
                                                <GridViewRowPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsSelected" Value="true">
                                                    <Setter TargetName="Bd" Property="Background" Value="$($c.PrimaryBg)"/>
                                                    <Setter Property="Foreground" Value="$($c.PrimaryFg)"/>
                                                </Trigger>
                                                <Trigger Property="IsMouseOver" Value="true">
                                                    <Setter TargetName="Bd" Property="Background" Value="$($c.HoverBg)"/>
                                                </Trigger>
                                                <MultiTrigger>
                                                    <MultiTrigger.Conditions>
                                                        <Condition Property="IsSelected" Value="true"/>
                                                        <Condition Property="IsMouseOver" Value="true"/>
                                                    </MultiTrigger.Conditions>
                                                    <Setter TargetName="Bd" Property="Background" Value="$($c.PrimaryBg)"/>
                                                    <Setter Property="Foreground" Value="$($c.PrimaryFg)"/>
                                                </MultiTrigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                                        <Setter Property="Background" Value="$($c.AltRowBg)"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </ListView.ItemContainerStyle>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Health" Width="70">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <TextBlock Text="{Binding HealthBadge}" Foreground="{Binding HealthColor}" FontWeight="SemiBold" FontSize="12" VerticalAlignment="Center"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header="Name"   DisplayMemberBinding="{Binding Name}"          Width="230"/>
                                <GridViewColumn Header="Driver" DisplayMemberBinding="{Binding DriverName}"    Width="220"/>
                                <GridViewColumn Header="Port"   DisplayMemberBinding="{Binding PortName}"      Width="140"/>
                                <GridViewColumn Header="Shared" DisplayMemberBinding="{Binding Shared}"        Width="55"/>
                                <GridViewColumn Header="Status" DisplayMemberBinding="{Binding PrinterStatus}" Width="95"/>
                            </GridView>
                        </ListView.View>
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="View Details"   x:Name="ctxPrnDetails"/>
                                <MenuItem Header="Rename Printer" x:Name="ctxPrnRename"/>
                                <MenuItem Header="Remove Printer" x:Name="ctxPrnRemove"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                    </ListView>
                </Border>
            </TabItem>

            <!-- Tab 1: Print Queue -->
            <TabItem Header="Print Queue" FontSize="13">
                <Border Background="$($c.Bg)" CornerRadius="0,6,6,6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <!-- Queue toolbar -->
                        <Border Grid.Row="0" Padding="10,8" Background="$($c.BtnBg)" BorderBrush="$($c.GridBorder)" BorderThickness="0,0,0,1">
                            <Grid>
                                <TextBlock x:Name="lblQueueStatus" Text="Refresh to load print jobs." Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                    <Button x:Name="btnRefreshQueue" Content="Refresh Queue" Width="120" Height="26" Margin="0,0,8,0" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                                    <Button x:Name="btnClearAllJobs" Content="Clear All Jobs" Width="110" Height="26" Background="$($c.BtnBg)" Foreground="$($c.Danger)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                        <!-- Queue list -->
                        <ListView x:Name="lvQueue" Grid.Row="1" BorderThickness="0" Margin="4" Background="Transparent" Foreground="$($c.Fg)" AlternationCount="2">
                            <ListView.ItemContainerStyle>
                                <Style TargetType="ListViewItem">
                                    <Setter Property="Height" Value="26"/>
                                    <Setter Property="Background" Value="Transparent"/>
                                    <Setter Property="Foreground" Value="$($c.Fg)"/>
                                    <Setter Property="Template">
                                        <Setter.Value>
                                            <ControlTemplate TargetType="ListViewItem">
                                                <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="0" CornerRadius="4" Padding="4,0" Margin="0,1">
                                                    <GridViewRowPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsSelected" Value="true">
                                                        <Setter TargetName="Bd" Property="Background" Value="$($c.PrimaryBg)"/>
                                                        <Setter Property="Foreground" Value="$($c.PrimaryFg)"/>
                                                    </Trigger>
                                                    <Trigger Property="IsMouseOver" Value="true">
                                                        <Setter TargetName="Bd" Property="Background" Value="$($c.HoverBg)"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Setter.Value>
                                    </Setter>
                                    <Style.Triggers>
                                        <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                                            <Setter Property="Background" Value="$($c.AltRowBg)"/>
                                        </Trigger>
                                    </Style.Triggers>
                                </Style>
                            </ListView.ItemContainerStyle>
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Printer"    DisplayMemberBinding="{Binding PrinterName}"  Width="200"/>
                                    <GridViewColumn Header="Document"   DisplayMemberBinding="{Binding DocumentName}" Width="200"/>
                                    <GridViewColumn Header="User"       DisplayMemberBinding="{Binding UserName}"     Width="120"/>
                                    <GridViewColumn Header="Pages"      DisplayMemberBinding="{Binding TotalPages}"   Width="60"/>
                                    <GridViewColumn Header="Size (KB)"  DisplayMemberBinding="{Binding SizeKB}"       Width="80"/>
                                    <GridViewColumn Header="Status"     DisplayMemberBinding="{Binding JobStatus}"    Width="100"/>
                                    <GridViewColumn Header="Submitted"  DisplayMemberBinding="{Binding SubmittedTime}" Width="140"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </Grid>
                </Border>
            </TabItem>

        </TabControl>

        <!-- Footer -->
        <Border Grid.Row="2" Background="$($c.BtnBg)" BorderBrush="$($c.GridBorder)" BorderThickness="0,1,0,0" Padding="12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnRefresh"    Content="Refresh"        Margin="0,0,8,0"  Width="90"  Height="28" Background="$($c.BtnBg)" Foreground="$($c.Fg)"     BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                <Button x:Name="btnPrnDetails" Content="View Details"   Margin="0,0,8,0"  Width="100" Height="28" Background="$($c.BtnBg)" Foreground="$($c.Fg)"     BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                <Button x:Name="btnPrnRemove"  Content="Remove Printer" Margin="0,0,8,0"  Width="110" Height="28" Background="$($c.BtnBg)" Foreground="$($c.Danger)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                <Button x:Name="btnClose"      Content="Close"                            Width="80"  Height="28" Background="$($c.BtnBg)" Foreground="$($c.Fg)"     BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
$pmReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($pmXaml))
$prnWin = [System.Windows.Markup.XamlReader]::Load($pmReader)

# Bind Controls
$lvPrinters        = $prnWin.FindName("lvPrinters")
$lvQueue           = $prnWin.FindName("lvQueue")
$lblHeaderTitle    = $prnWin.FindName("lblHeaderTitle")
$lblDescription    = $prnWin.FindName("lblDescription")
$lblStatus         = $prnWin.FindName("lblStatus")
$lblQueueStatus    = $prnWin.FindName("lblQueueStatus")
$btnRestartSpooler = $prnWin.FindName("btnRestartSpooler")
$btnInstallPrinter = $prnWin.FindName("btnInstallPrinter")
$btnRefresh        = $prnWin.FindName("btnRefresh")
$btnRefreshQueue   = $prnWin.FindName("btnRefreshQueue")
$btnClearAllJobs   = $prnWin.FindName("btnClearAllJobs")
$btnPrnDetails     = $prnWin.FindName("btnPrnDetails")
$btnPrnRemove      = $prnWin.FindName("btnPrnRemove")
$btnClose          = $prnWin.FindName("btnClose")

$ctxPrnDetails = $prnWin.FindName("ctxPrnDetails")
$ctxPrnRename  = $prnWin.FindName("ctxPrnRename")
$ctxPrnRemove  = $prnWin.FindName("ctxPrnRemove")

# 6. Primary Action Logic

$RefreshPrinters = {
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    $lblStatus.Text = "Querying printers from $ComputerName..."
    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Orange

    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)

    # Query OS description once
    if ($lblDescription -and $lblDescription.Text -eq "Asset: Querying...") {
        try {
            $desc = Invoke-Command -ComputerName $ComputerName -ScriptBlock { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Description } -ErrorAction SilentlyContinue
            $lblDescription.Text = "Asset: " + (if ([string]::IsNullOrWhiteSpace($desc)) { "Not Set" } else { $desc })
        } catch { $lblDescription.Text = "Asset: Unavailable" }
    }

    try {
        $printers = @(Get-RemotePrinters -ComputerName $ComputerName)

        # --- Printer Health Check ---
        # Extract IP from PortName (common formats: "IP_10.1.2.3", "10.1.2.3", "TCP-10.1.2.3")
        # Run checks in parallel jobs, one per unique IP, then map results back.
        $ipMap    = @{}   # portName -> IP string
        $ipsToCheck = @()
        foreach ($p in $printers) {
            $port = $p.PortName
            $ip   = $null
            if ($port -match '(\d{1,3}(?:\.\d{1,3}){3})') { $ip = $matches[1] }
            if ($ip) { $ipMap[$port] = $ip; if ($ip -notin $ipsToCheck) { $ipsToCheck += $ip } }
        }

        # Run health checks concurrently via jobs
        $healthJobs = @{}
        foreach ($ip in $ipsToCheck) {
            $healthJobs[$ip] = Start-Job -ScriptBlock {
                param($addr)
                $ping = $false; $port9100 = $false
                try {
                    $p = New-Object System.Net.NetworkInformation.Ping
                    if ($p.Send($addr, 1200).Status -eq 'Success') { $ping = $true }
                } catch { Write-Debug "Ping failed for ${addr}: $($_.Exception.Message)" }
                try {
                    $tcp = New-Object System.Net.Sockets.TcpClient
                    $ar  = $tcp.BeginConnect($addr, 9100, $null, $null)
                    if ($ar.AsyncWaitHandle.WaitOne(1200, $false) -and $tcp.Connected) { $port9100 = $true }
                    $tcp.Close()
                } catch { Write-Debug "TCP 9100 check failed for ${addr}: $($_.Exception.Message)" }
                return [PSCustomObject]@{ IP = $addr; Ping = $ping; Port9100 = $port9100 }
            } -ArgumentList $ip
        }

        # Wait up to 4 s for all jobs (most LAN pings resolve in < 500 ms)
        $deadline = (Get-Date).AddSeconds(4)
        while ($healthJobs.Values | Where-Object { $_.State -eq 'Running' }) {
            if ((Get-Date) -gt $deadline) { break }
            Start-Sleep -Milliseconds 150
        }

        # Collect results
        $healthResults = @{}
        foreach ($ip in $healthJobs.Keys) {
            try {
                $r = Receive-Job $healthJobs[$ip] -ErrorAction SilentlyContinue
                if ($r) { $healthResults[$ip] = $r }
            } catch { Write-Warning "Failed to receive job for ${ip}: $($_.Exception.Message)" }
            Stop-Job  $healthJobs[$ip] -ErrorAction SilentlyContinue
            Remove-Job $healthJobs[$ip] -Force -ErrorAction SilentlyContinue
        }

        # Stamp each printer with badge + color
        foreach ($p in $printers) {
            $ip  = $ipMap[$p.PortName]
            $r   = if ($ip) { $healthResults[$ip] } else { $null }

            if (-not $ip) {
                # Local/USB printer or non-IP port -- no health data
                $p | Add-Member -MemberType NoteProperty -Name "HealthBadge" -Value "N/A"  -Force
                $p | Add-Member -MemberType NoteProperty -Name "HealthColor" -Value $c.SecFg -Force
            } elseif ($r -and $r.Port9100) {
                $p | Add-Member -MemberType NoteProperty -Name "HealthBadge" -Value "Online" -Force
                $p | Add-Member -MemberType NoteProperty -Name "HealthColor" -Value "#22C55E" -Force
            } elseif ($r -and $r.Ping) {
                $p | Add-Member -MemberType NoteProperty -Name "HealthBadge" -Value "Ping OK" -Force
                $p | Add-Member -MemberType NoteProperty -Name "HealthColor" -Value "#F59E0B" -Force
            } else {
                $p | Add-Member -MemberType NoteProperty -Name "HealthBadge" -Value "Offline" -Force
                $p | Add-Member -MemberType NoteProperty -Name "HealthColor" -Value $c.Danger -Force
            }
        }

        if ($lvPrinters) { $lvPrinters.ItemsSource = @($printers | Sort-Object Name) }
        $lblStatus.Text = "Found $($printers.Count) printers. (Updated: $(Get-Date -Format 'HH:mm:ss'))"
        $lblStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#22C55E"))
    } catch {
        $lblStatus.Text = "Error communicating with $ComputerName."
        $lblStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Show-LocalMessageBox -Message "Failed to fetch printers:`n$($_.Exception.Message)" -Title "Connection Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
    }
    [System.Windows.Input.Mouse]::OverrideCursor = $null
}

$RefreshQueue = {
    if (-not $lvQueue) { return }
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    if ($lblQueueStatus) { $lblQueueStatus.Text = "Querying print jobs from $ComputerName..." }

    try {
        $jobs = @(Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-PrintJob -PrinterName * -ErrorAction SilentlyContinue |
                Select-Object Id, PrinterName, DocumentName, UserName,
                              TotalPages, PagesPrinted, JobStatus,
                              @{ Name='SizeKB'; Expression={ [math]::Round($_.Size / 1KB, 1) } },
                              SubmittedTime
        } -ErrorAction Stop)

        $lvQueue.ItemsSource = @($jobs | Sort-Object SubmittedTime -Descending)
        $statusText = if ($jobs.Count -eq 0) { "Print queue is empty." } else { "$($jobs.Count) job(s) in queue. (Updated: $(Get-Date -Format 'HH:mm:ss'))" }
        if ($lblQueueStatus) { $lblQueueStatus.Text = $statusText }
    } catch {
        if ($lblQueueStatus) { $lblQueueStatus.Text = "Failed to query print jobs: $($_.Exception.Message)" }
        $lvQueue.ItemsSource = @()
    }
    [System.Windows.Input.Mouse]::OverrideCursor = $null
}

$ShowPrinterDetails = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $selP = $lvPrinters.SelectedItem

        # Embedded Details XAML -- card row layout matching Device Info tab style
        $detXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Printer Details" Width="520" SizeToContent="Height" WindowStartupLocation="CenterOwner"
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
                                    <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.4"/></Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </Window.Resources>
            <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
                <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3" Color="Black"/></Border.Effect>
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="20,20,20,0" Cursor="Hand">
                        <StackPanel>
                            <TextBlock Text="Printer Details" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                            <TextBlock Text="$($selP.Name)" FontSize="12" Foreground="$($c.SecFg)" Margin="0,3,0,0"/>
                        </StackPanel>
                    </Border>
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="20,14,20,20">
                        <StackPanel>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.BtnBg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Printer Name" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Name" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.Bg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Driver" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Driver" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.BtnBg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Port / IP" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Port" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.Bg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Shared" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Shared" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.BtnBg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Published" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Published" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.Bg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Status" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Status" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,6" Padding="14,10" Background="$($c.BtnBg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Location" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Location" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                            </Border>
                            <Border Margin="0,0,0,0" Padding="14,10" Background="$($c.Bg)" CornerRadius="6" BorderBrush="$($c.GridBorder)" BorderThickness="1">
                                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Comment" FontWeight="SemiBold" Foreground="$($c.SecFg)" FontSize="13" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtDet_Comment" Grid.Column="1" Text="-" Foreground="$($c.Fg)" FontSize="13" VerticalAlignment="Center" TextWrapping="Wrap"/></Grid>
                            </Border>
                        </StackPanel>
                    </ScrollViewer>
                    <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                        <Button x:Name="btnDetClose" Content="Close" HorizontalAlignment="Right" Width="80" Height="28" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </Border>
                </Grid>
            </Border>
        </Window>
"@
        $dReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($detXaml))
        $detWin = [System.Windows.Markup.XamlReader]::Load($dReader)
        $detWin.Owner = $prnWin
        $detWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $detWin.DragMove() })

        # Populate individual card fields
        $detWin.FindName("txtDet_Name").Text      = if ($selP.Name)          { $selP.Name }          else { "-" }
        $detWin.FindName("txtDet_Driver").Text    = if ($selP.DriverName)    { $selP.DriverName }    else { "-" }
        $detWin.FindName("txtDet_Port").Text      = if ($selP.PortName)      { $selP.PortName }      else { "-" }
        $detWin.FindName("txtDet_Shared").Text    = if ($null -ne $selP.Shared)    { "$($selP.Shared)" }    else { "-" }
        $detWin.FindName("txtDet_Published").Text = if ($null -ne $selP.Published) { "$($selP.Published)" } else { "-" }
        $detWin.FindName("txtDet_Status").Text    = if ($selP.PrinterStatus) { "$($selP.PrinterStatus)" } else { "-" }
        $detWin.FindName("txtDet_Location").Text  = if ($selP.Location)      { $selP.Location }      else { "-" }
        $detWin.FindName("txtDet_Comment").Text   = if ($selP.Comment)       { $selP.Comment }       else { "-" }

        $detWin.FindName("btnDetClose").Add_Click({ $detWin.Close() })
        Center-OnPrnWin -ChildWindow $detWin -OwnerWindow $prnWin
        $detWin.ShowDialog() | Out-Null
    }
}

$RemovePrinterAction = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $pName = $lvPrinters.SelectedItem.Name
        $conf = Show-LocalMessageBox -Message "Are you sure you want to permanently remove the printer '$pName' from $ComputerName?" -Title "Confirm Removal" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $prnWin
        
        if ($conf -eq "Yes") {
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            try {
                Remove-RemotePrinter -ComputerName $ComputerName -PrinterName $pName
                Add-AppLog -Event "Printer Remove" -Username "System" -Details "Removed printer '$pName' from $ComputerName." -Config $Config -State $State -Status "Success"
                Show-LocalMessageBox -Message "Printer '$pName' removed successfully." -Title "Success" -OwnerWindow $prnWin | Out-Null
                & $RefreshPrinters
            } catch { 
                Show-LocalMessageBox -Message "Failed to remove printer:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
                Add-AppLog -Event "Printer Remove" -Username "System" -Details "Failed to remove '$pName' on $($ComputerName): $($_.Exception.Message)" -Config $Config -State $State -Status "Error"
            }
            [System.Windows.Input.Mouse]::OverrideCursor = $null
        }
    }
}

$RenamePrinterAction = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $pName = $lvPrinters.SelectedItem.Name
        
        # Embedded Rename Dialog XAML
        $renXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Rename Printer" Width="380" SizeToContent="Height" WindowStartupLocation="CenterOwner" 
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
            <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
                <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand">
                        <TextBlock Text="Rename Printer" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                    </Border>
                    
                    <StackPanel Grid.Row="1" Margin="16,8,16,16">
                        <TextBlock Text="Enter a new name for '$pName':" FontSize="12" Foreground="$($c.SecFg)" Margin="0,0,0,4" TextWrapping="Wrap"/>
                        <TextBox x:Name="txtNewName" Height="30" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Padding="6,4" VerticalContentAlignment="Center" Text="$pName"/>
                    </StackPanel>
                    
                    <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnRenOk" Content="Rename" Width="80" Height="28" Margin="0,0,8,0" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0" IsDefault="True"/>
                            <Button x:Name="btnRenCancel" Content="Cancel" Width="80" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" IsCancel="True"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </Border>
        </Window>
"@
        $rReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($renXaml))
        $renWin = [System.Windows.Markup.XamlReader]::Load($rReader)
        $renWin.Owner = $prnWin
        
        $renWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $renWin.DragMove() })
        
        $txtNewName = $renWin.FindName("txtNewName")
        $btnRenOk = $renWin.FindName("btnRenOk")
        $btnRenCancel = $renWin.FindName("btnRenCancel")
        
        $script:newPrnName = ""
        
        if ($btnRenCancel) { $btnRenCancel.Add_Click({ $renWin.Close() }) }
        if ($btnRenOk) {
            $btnRenOk.Add_Click({
                if (-not [string]::IsNullOrWhiteSpace($txtNewName.Text)) {
                    $script:newPrnName = $txtNewName.Text
                    $renWin.Close()
                } else {
                    Show-LocalMessageBox -Message "Printer name cannot be blank." -Title "Validation" -IconType "Warning" -OwnerWindow $renWin | Out-Null
                }
            })
        }
        
        $txtNewName.SelectAll()
        $txtNewName.Focus() | Out-Null
        Center-OnPrnWin -ChildWindow $renWin -OwnerWindow $prnWin
        $renWin.ShowDialog() | Out-Null
        
        if ($script:newPrnName -and $script:newPrnName -ne $pName) {
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            try {
                Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                    param($old, $new)
                    Rename-Printer -Name $old -NewName $new -ErrorAction Stop
                } -ArgumentList $pName, $script:newPrnName
                
                Add-AppLog -Event "Printer Rename" -Username "System" -Details "Renamed printer '$pName' to '$($script:newPrnName)' on $ComputerName." -Config $Config -State $State -Status "Success"
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                Show-LocalMessageBox -Message "Printer renamed to '$($script:newPrnName)' successfully." -Title "Success" -IconType "Information" -OwnerWindow $prnWin | Out-Null
                & $RefreshPrinters
            } catch {
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                Show-LocalMessageBox -Message "Failed to rename printer:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
                Add-AppLog -Event "Printer Rename" -Username "System" -Details "Failed to rename '$pName' on $($ComputerName): $($_.Exception.Message)" -Config $Config -State $State -Status "Error"
            }
        }
    }
}

$InstallPrinterAction = {
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    $lblStatus.Text = "Fetching remote driver manifest (15s timeout)..."
    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Orange
    
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    
    $remoteDrivers = @()
    $job = $null
    try {
        $job = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-PrinterDriver | Select-Object -ExpandProperty Name } -AsJob
        
        if ($job -and (Wait-Job $job -Timeout 15)) {
            if ($job.State -eq 'Failed') { throw $job.ChildJobs[0].JobStateInfo.Reason }
            $remoteDrivers = @(Receive-Job $job -ErrorAction Stop | Sort-Object)
        } else {
            if ($job) { Stop-Job $job }
            throw "Connection timed out (15s)."
        }
    } catch {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $lblStatus.Text = "Driver fetch failed."
        Show-LocalMessageBox -Message "Unable to retrieve drivers from ${ComputerName}.`n`nError: $($_.Exception.Message)" -Title "Fetch Failed" -IconType "Warning" -OwnerWindow $prnWin | Out-Null
        return
    } finally {
        if ($job) { Remove-Job $job -ErrorAction SilentlyContinue }
    }
    [System.Windows.Input.Mouse]::OverrideCursor = $null
    $lblStatus.Text = "Ready."
    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Green
    
    # Embedded Install Dialog XAML
    $instXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Remote Printer Install" Width="420" SizeToContent="Height" WindowStartupLocation="CenterOwner" 
            WindowStyle="None" AllowsTransparency="True" Background="Transparent"
            FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
        <Window.Resources>
            <Style TargetType="Button">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style TargetType="TextBox">
                <Setter Property="Background" Value="$($c.BtnBg)"/>
                <Setter Property="Foreground" Value="$($c.Fg)"/>
                <Setter Property="BorderBrush" Value="$($c.BtnBorder)"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="Padding" Value="6,4"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="TextBox">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ScrollViewer x:Name="PART_ContentHost"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Window.Resources>
        <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand">
                    <TextBlock Text="Install Printer on $ComputerName" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                </Border>
                
                <StackPanel Grid.Row="1" Margin="16,8,16,16">
                    <TextBlock Text="Printer Name (Friendly Name):" FontSize="11" Foreground="$($c.SecFg)" Margin="0,0,0,4"/>
                    <TextBox x:Name="txtName" Height="30" Margin="0,0,0,12"/>
                    
                    <TextBlock Text="Driver Name (Select or Type):" FontSize="11" Foreground="$($c.SecFg)" Margin="0,0,0,4"/>
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="32"/>
                            <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <ComboBox x:Name="cbDriver" Grid.Column="0" Height="30" IsEditable="True" Margin="0,0,4,0"/>
                        <Button x:Name="btnLocalDriver" Grid.Column="1" Grid.ColumnSpan="2" Content="Local Driver" FontSize="11" FontWeight="Bold" ToolTip="Copy installed driver from this PC" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </Grid>
                    
                    <TextBlock Text="IP Address:" FontSize="11" Foreground="$($c.SecFg)" Margin="0,0,0,4"/>
                    <TextBox x:Name="txtIP" Height="30"/>
                </StackPanel>
                
                <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="btnInstall" Content="Install" Width="80" Height="28" Margin="0,0,8,0" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0" IsDefault="True"/>
                        <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" IsCancel="True" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
    </Window>
"@
    $iReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($instXaml))
    $instWin = [System.Windows.Markup.XamlReader]::Load($iReader)
    
    $instWin.Owner = $prnWin
    $instWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $instWin.DragMove() })
    
    $cbDriver = $instWin.FindName("cbDriver")
    if ($cbDriver -and $remoteDrivers) { 
        $cbDriver.ItemsSource = $remoteDrivers
        if ($cbDriver.Items.Count -gt 0) { $cbDriver.SelectedIndex = 0 }
    }
    
    $txtName = $instWin.FindName("txtName")
    $txtIP = $instWin.FindName("txtIP")
    $btnUploadDriver = $instWin.FindName("btnUploadDriver")
    $btnLocalDriver = $instWin.FindName("btnLocalDriver")
    $btnInstall = $instWin.FindName("btnInstall")
    $btnCancel = $instWin.FindName("btnCancel")

    $script:prnInput = $null

    if ($btnUploadDriver) {
        $btnUploadDriver.Add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = "Driver Information File (*.inf)|*.inf"
            $ofd.Title = "Select Driver INF File"
            
            if ($ofd.ShowDialog() -eq "OK") {
                try {
                    $infPath = $ofd.FileName
                    
                    # Wrap heavy operations in the new progress window
                    $updatedRemoteDrivers = Invoke-WithProgress -Title "Staging Driver File" -OwnerWindow $instWin -Action {
                        param($Update)
                        
                        &$Update "Analyzing driver file structure..." 10
                        $drvDir = Split-Path $infPath
                        $folderName = Split-Path $drvDir -Leaf
                        $remoteTemp = "\\$ComputerName\c$\Temp\HDC_Drivers"
                        
                        &$Update "Initializing remote staging directory..." 30
                        if (-not (Test-Path $remoteTemp)) { New-Item -ItemType Directory -Path $remoteTemp -Force | Out-Null }
                        
                        &$Update "Copying files over network (this may take a minute depending on driver size)..." 50
                        Copy-Item -Path $drvDir -Destination $remoteTemp -Recurse -Force
                        $remoteLocalPath = "C:\Temp\HDC_Drivers\$folderName"
                        
                        &$Update "Executing native PnPUtil to install driver remotely..." 80
                        $res = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                            param($path)
                            $res = pnputil.exe /add-driver "$path\*.inf" /subdirs /install
                            $drvList = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object)
                            return [PSCustomObject]@{ Pnp = ($res | Out-String); Drivers = $drvList }
                        } -ArgumentList $remoteLocalPath
                        
                        &$Update "Finalizing driver list..." 100
                        return $res
                    }
                    
                    # Refresh the combo box with remote drivers
                    if ($cbDriver -and $updatedRemoteDrivers.Drivers) {
                        $currentText = $cbDriver.Text
                        $cbDriver.ItemsSource = $updatedRemoteDrivers.Drivers
                        $cbDriver.Text = $currentText
                    }
                    
                    Show-LocalMessageBox -Message "Driver staged successfully.`n`nOutput:`n$($updatedRemoteDrivers.Pnp)`n`nPlease select the new driver from the dropdown or type its exact Model Name." -Title "Driver Staged" -IconType "Information" -OwnerWindow $instWin | Out-Null
                } catch { 
                    Show-LocalMessageBox -Message "Failed to deploy driver:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $instWin | Out-Null 
                }
            }
        })
    }
    
    if ($btnLocalDriver) {
        $btnLocalDriver.Add_Click({
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            Import-Module PrintManagement -ErrorAction SilentlyContinue
            
            $localDrivers = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object Name, InfPath | Sort-Object Name)
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            
            if (-not $localDrivers -or $localDrivers.Count -eq 0) {
                Show-LocalMessageBox -Message "No local drivers were found on this computer. Ensure you run as Administrator." -Title "No Drivers" -IconType "Warning" -OwnerWindow $instWin | Out-Null
                return
            }

            # Embedded Select Driver XAML
            $selXaml = @"
            <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                    Title="Select Local Driver" Width="400" Height="500" WindowStartupLocation="CenterOwner" 
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
                <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
                    <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand">
                            <TextBlock Text="Select Local Driver" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                        </Border>
                        
                        <ListBox x:Name="lbDrivers" Grid.Row="1" Margin="16,8,16,16" DisplayMemberPath="Name" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                        
                        <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button x:Name="btnSelOk" Content="Use Selected" Width="100" Height="28" Margin="0,0,8,0" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0" IsDefault="True"/>
                                <Button x:Name="btnSelCancel" Content="Cancel" Width="80" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" IsCancel="True"/>
                            </StackPanel>
                        </Border>
                    </Grid>
                </Border>
            </Window>
"@
            $sReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($selXaml))
            $locWin = [System.Windows.Markup.XamlReader]::Load($sReader)
            
            $locWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $locWin.DragMove() })
            
            $lbLoc = $locWin.FindName("lbDrivers")
            $btnSelOk = $locWin.FindName("btnSelOk")
            $btnSelCancel = $locWin.FindName("btnSelCancel")
            
            $lbLoc.ItemsSource = $localDrivers
            $locWin.Owner = $instWin 
            
            $script:selectedLocalDriver = $null
            
            if ($btnSelCancel) { $btnSelCancel.Add_Click({ $locWin.Close() }) }
            if ($btnSelOk) {
                $btnSelOk.Add_Click({
                    if ($lbLoc.SelectedItem) {
                        $script:selectedLocalDriver = $lbLoc.SelectedItem
                        $locWin.Close()
                    }
                })
            }
            
            Center-OnPrnWin -ChildWindow $locWin -OwnerWindow $prnWin
            $locWin.ShowDialog() | Out-Null

            if ($script:selectedLocalDriver) {
                try {
                    $drvName = $script:selectedLocalDriver.Name
                    $infPath = $script:selectedLocalDriver.InfPath
                    
                    if (-not $infPath -or -not (Test-Path $infPath)) {
                        throw "Could not locate the physical INF file directory for driver: $drvName"
                    }
                    
                    # Wrapping deployment in progress dialog
                    $updatedRemoteDrivers = Invoke-WithProgress -Title "Deploying Local Driver" -OwnerWindow $instWin -Action {
                        param($Update)
                        
                        &$Update "Preparing driver architecture files..." 10
                        $drvDir = Split-Path $infPath
                        $folderName = Split-Path $drvDir -Leaf
                        $remoteTemp = "\\$ComputerName\c$\Temp\HDC_Drivers"
                        
                        &$Update "Initializing remote staging directory..." 30
                        if (-not (Test-Path $remoteTemp)) { New-Item -ItemType Directory -Path $remoteTemp -Force | Out-Null }
                        
                        &$Update "Copying driver structure from Local C: to Remote C: (this may take a minute depending on driver size)..." 50
                        Copy-Item -Path $drvDir -Destination $remoteTemp -Recurse -Force
                        $remoteLocalPath = "C:\Temp\HDC_Drivers\$folderName"
                        
                        &$Update "Executing native PnPUtil to install driver remotely..." 80
                        $res = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                            param($path, $name)
                            pnputil.exe /add-driver "$path\*.inf" /subdirs /install | Out-Null
                            
                            try { Add-PrinterDriver -Name $name -ErrorAction Stop } catch { Write-Warning $_.Exception.Message }
                            return @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object)
                        } -ArgumentList $remoteLocalPath, $drvName
                        
                        &$Update "Finishing deployment..." 100
                        return $res
                    }
                    
                    # Guarantee it appears in the Combobox
                    if ($cbDriver -and $updatedRemoteDrivers) {
                        $cbDriver.ItemsSource = $updatedRemoteDrivers
                        
                        if (-not $cbDriver.Items.Contains($drvName)) {
                            # If it's still not there (unexpected), we must add it to the underlying source
                            # since Items.Add() is prohibited when ItemsSource is set.
                            $updatedRemoteDrivers += $drvName
                            $cbDriver.ItemsSource = $updatedRemoteDrivers
                        }

                        $cbDriver.SelectedItem = $drvName
                        $cbDriver.Text = $drvName 
                        $cbDriver.Focus()
                    }
                    Show-LocalMessageBox -Message "Driver '$drvName' uploaded, staged, and natively loaded into the remote spooler successfully." -Title "Success" -IconType "Information" -OwnerWindow $instWin | Out-Null
                } catch { 
                    Show-LocalMessageBox -Message "Deploy failed:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $instWin | Out-Null 
                }
            }
        })
    }
    
    if ($btnCancel) { $btnCancel.Add_Click({ $instWin.Close() }) }
    
    # Restored proper Install submission format with Validation
    if ($btnInstall) {
        $btnInstall.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtName.Text) -or [string]::IsNullOrWhiteSpace($cbDriver.Text)) {
                Show-LocalMessageBox -Message "A Printer Name and Driver Name are required to begin installation." -Title "Validation" -IconType "Warning" -OwnerWindow $instWin | Out-Null
                return
            }
            
            # Map valid inputs before closing
            $script:prnInput = @{
                Name = $txtName.Text
                Driver = $cbDriver.Text
                IP = $txtIP.Text
            }
            $instWin.Close()
        })
    }
    
    Center-OnPrnWin -ChildWindow $instWin -OwnerWindow $prnWin
    $instWin.ShowDialog() | Out-Null
    
    if ($script:prnInput) {
        Add-AppLog -Event "Printer Install" -Username "System" -Details "Installing '$($script:prnInput.Name)' on $ComputerName..." -Config $Config -State $State -Status "Info"
        
        try {
            # Wrapped final deployment in progress window for UX consistency
            Invoke-WithProgress -Title "Installing Printer" -OwnerWindow $prnWin -Action {
                param($Update)
                &$Update "Building printer object and registering '$($script:prnInput.Name)' on $ComputerName..." 20
                Install-RemotePrinter -ComputerName $ComputerName -PrinterName $script:prnInput.Name -DriverName $script:prnInput.Driver -IPAddress $script:prnInput.IP
                &$Update "Installation complete." 100
            }

            Show-LocalMessageBox -Message "Printer installed successfully on ${ComputerName}." -Title "Success" -IconType "Information" -OwnerWindow $prnWin | Out-Null
            Add-AppLog -Event "Printer Install" -Username "System" -Details "Successfully installed printer on ${ComputerName}." -Config $Config -State $State -Status "Success" -Color "Green"
            & $RefreshPrinters
        } catch {
            Show-LocalMessageBox -Message "Installation failed:`n$($_.Exception.Message)" -Title "Remote Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
            Add-AppLog -Event "Printer Install" -Username "System" -Details "Failed on ${ComputerName}: $($_.Exception.Message)" -Config $Config -State $State -Status "Error" -Color "Red"
        }
    }
}

$RestartSpoolerAction = {
    $conf = Show-LocalMessageBox -Message "Are you sure you want to restart the Print Spooler on $ComputerName?" -Title "Confirm Restart" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $prnWin
    if ($conf -eq "Yes") {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        try {
            Restart-RemoteSpooler -ComputerName $ComputerName
            Add-AppLog -Event "Service" -Username "System" -Details "Spooler restarted on $ComputerName" -Config $Config -State $State
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-LocalMessageBox -Message "Print Spooler restarted successfully." -Title "Success" -OwnerWindow $prnWin | Out-Null
            & $RefreshPrinters
        } catch {
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-LocalMessageBox -Message "Failed to restart spooler:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
        }
    }
}

# 7. Wire Up Bindings
if ($btnRefresh) { $btnRefresh.Add_Click($RefreshPrinters) }
if ($btnClose)   { $btnClose.Add_Click({ $prnWin.Close() }) }

if ($btnPrnDetails) { $btnPrnDetails.Add_Click($ShowPrinterDetails) }
if ($ctxPrnDetails) { $ctxPrnDetails.Add_Click($ShowPrinterDetails) }
if ($ctxPrnRename)  { $ctxPrnRename.Add_Click($RenamePrinterAction) }
if ($lvPrinters)    { $lvPrinters.Add_MouseDoubleClick($ShowPrinterDetails) }

if ($btnPrnRemove) { $btnPrnRemove.Add_Click($RemovePrinterAction) }
if ($ctxPrnRemove) { $ctxPrnRemove.Add_Click($RemovePrinterAction) }

if ($btnInstallPrinter) { $btnInstallPrinter.Add_Click($InstallPrinterAction) }
if ($btnRestartSpooler) { $btnRestartSpooler.Add_Click($RestartSpoolerAction) }

# Print Queue wiring
if ($btnRefreshQueue) { $btnRefreshQueue.Add_Click($RefreshQueue) }
if ($btnClearAllJobs) {
    $btnClearAllJobs.Add_Click({
        $conf = Show-LocalMessageBox -Message "Clear ALL pending print jobs on $ComputerName?`nThis cannot be undone." -Title "Confirm Clear" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $prnWin
        if ($conf -eq "Yes") {
            try {
                Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                    Get-PrintJob -PrinterName * -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
                } -ErrorAction Stop
                & $RefreshQueue
                Show-LocalMessageBox -Message "All print jobs cleared." -Title "Done" -OwnerWindow $prnWin | Out-Null
            } catch {
                Show-LocalMessageBox -Message "Failed to clear jobs:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
            }
        }
    }.GetNewClosure())
}

# Keyboard navigation on the main printer manager window
$prnWin.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq 'F5')     { $e.Handled = $true; & $RefreshPrinters }
    if ($e.Key -eq 'Escape') { $e.Handled = $true; $prnWin.Close() }
    if ($e.Key -eq 'Delete' -and $lvPrinters -and $lvPrinters.IsKeyboardFocusWithin -and $lvPrinters.SelectedItem) {
        $e.Handled = $true; & $RemovePrinterAction
    }
}.GetNewClosure())

# Apply multi-monitor centering to all child windows via closure-based wrappers
# We patch the existing action scriptblocks to call Center-OnPrnWin before ShowDialog.
# Details window centering -- wrap ShowPrinterDetails
$OrigShowPrinterDetails = $ShowPrinterDetails
$ShowPrinterDetails = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $selP = $lvPrinters.SelectedItem
        # Build detWin exactly as before (reuse original logic)
        & $OrigShowPrinterDetails
    }
}.GetNewClosure()
# Note: centering of detWin, renWin, instWin, locWin is applied inline below via
# a post-load hook on each window using Add_Loaded.

# 8. Start
$prnWin.Add_Loaded({
    & $RefreshPrinters

    # Hook: center detWin when it loads
    # We intercept at the ShowDialog call sites using a small DispatcherTimer trick:
    # each child window registers its own Add_Loaded to call Center-OnPrnWin.
    # This is handled inside each action block (see ShowPrinterDetails, etc.).
})
$prnWin.ShowDialog() | Out-Null
