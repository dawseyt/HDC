# ============================================================================
# Main.ps1 - Modular Account Monitor Entry Point
# ============================================================================

<#
.SYNOPSIS
    Lightweight main controller for the modularized Account Monitoring tool.
.DESCRIPTION
    Loads the backend PowerShell modules, parses the main MainWindow.xaml, 
    initializes app state, and routes UI events to the UIEvents module.
#>

# --- 1. ENVIRONMENT SETUP ---
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    
    # Enforce TLS 1.2 for APIs
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
catch {
    [System.Windows.MessageBox]::Show("Failed to load required .NET assemblies. $($_.Exception.Message)", "Initialization Error", "OK", "Error")
    exit
}

# --- 2. IMPORT MODULES ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Warning "Active Directory PowerShell module is required. Some functionality may be limited."
}

# Import our custom backend modules
Import-Module ".\modules\CoreLogic.psm1" -Force -DisableNameChecking
Import-Module ".\modules\ActiveDirectory.psm1" -Force -DisableNameChecking
Import-Module ".\modules\RemoteManagement.psm1" -Force -DisableNameChecking
Import-Module ".\modules\UIEvents.psm1" -Force -DisableNameChecking

# --- 3. INITIALIZE CONFIG AND STATE ---
$global:Config = Get-AppConfig

$global:State = @{
    UIControls             = @{}
    Actions                = @{}   # Cross-module scriptblock dispatch
    Timer                  = $null
    AutoUnlockTimer        = $null
    CurrentTheme           = $global:Config.GeneralSettings.DefaultTheme
    LoggedLockouts         = @{}
    LastSortCol            = $null
    SortDescending         = $false
    IsSearchPaused         = $false
    IsRefreshing           = $false # Guard for async RefreshAction
    RefreshTargetTime      = (Get-Date)
    RefreshIntervalSeconds = 180    # Default 3 mins
    SearchHistory          = @()
    CurrentTabComputer     = $null
}

Initialize-LogDirectory -Config $global:Config

# --- 4. LOAD MAIN UI (XAML) ---
$xamlPath = Join-Path $PSScriptRoot "ui\MainWindow.xaml"
if (-not (Test-Path $xamlPath)) {
    [System.Windows.MessageBox]::Show("Cannot find the UI definition file!`nExpected location: $xamlPath", "Missing UI File", "OK", "Error")
    exit
}

try {
    # Read the file and strip any markdown backticks that may have been accidentally copied/saved
    $xamlContent = [System.IO.File]::ReadAllText($xamlPath)
    $xamlContent = $xamlContent -replace '(?s)^```(xml|xaml)?\s*', ''
    $xamlContent = $xamlContent -replace '(?s)\s*```\s*$', ''
    $xamlContent = $xamlContent.Trim()

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlContent))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.MessageBox]::Show("Failed to parse MainWindow.xaml.`n`nError: $($_.Exception.Message)`n`nPlease ensure the UI file contains valid XML and no markdown artifacts.", "XAML Parsing Error", "OK", "Error")
    exit
}

# --- 5. INITIALIZE UI ELEMENTS FROM CONFIG ---
# Populate auto-refresh dropdown
$cbAutoRefresh = $window.FindName("cbAutoRefresh")
$cbAutoRefresh.ItemsSource = $global:Config.AutoSettings.AutoRefreshOptions | Where-Object { $_ -ne "Off" }
$cbAutoRefresh.SelectedIndex = 3 
$window.FindName("chkEnableEmail").IsChecked = $global:Config.EmailSettings.EnableEmailNotifications

# Apply UI Text from Config
if ($global:Config.ControlProperties) {
    $window.FindName("lblTitle").Text = $global:Config.ControlProperties.TitleLabel.Text
    $window.FindName("lblSubtitle").Text = $global:Config.ControlProperties.SubtitleLabel.Text
    $window.FindName("btnRefresh").Content = $global:Config.ControlProperties.RefreshButton.Text
    $window.FindName("btnUnlock").Content = $global:Config.ControlProperties.UnlockButton.Text
    $window.FindName("btnUnlockAll").Content = $global:Config.ControlProperties.UnlockAllButton.Text
    $window.FindName("btnSearch").Content = $global:Config.ControlProperties.SearchButton.Text
    $window.FindName("btnViewLog").Content = $global:Config.ControlProperties.ViewLogButton.Text
}

# Load Icon
$iconPath = Join-Path $PSScriptRoot "ADLockoutMgrIco.jpg"
if (Test-Path $iconPath) {
    try {
        $iconUri = [Uri]::new($iconPath)
        $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bitmap.BeginInit()
        $bitmap.UriSource = $iconUri
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.EndInit()
        $window.Icon = $bitmap
    } catch {
        Write-Warning "Failed to load icon: $_"
    }
}

# --- 6. REGISTER EVENTS & START ---
# This delegates all UI interactivity to the UIEvents module you have in the Canvas
Register-MainUIEvents -Window $window -Config $global:Config -State $global:State

Add-AppLog -Event "System" -Username "System" -Details "Application Started (Modular Architecture)" -Config $global:Config -State $global:State -Status "Info" -Color "Blue"
Add-AppLog -Event "Config" -Username "System" -Details "Using configuration from: $($global:CurrentConfigPath)" -Config $global:Config -State $global:State -Status "Info" -Color "Black"

# Note: The 'ApplyTheme' function is invoked securely inside the 'Window.Loaded' event mapped in Register-MainUIEvents
$window.ShowDialog() | Out-Null