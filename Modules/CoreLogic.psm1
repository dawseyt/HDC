# ============================================================================
# CoreLogic.psm1 - Configuration, Logging, and App Setup
# ============================================================================

# ---------------------------------------------------------------------------
# Credential Manager helpers
# Reads/writes the Freshservice API key from Windows Credential Manager
# (target name: "HDCompanion_FreshserviceAPIKey") so the key is never stored
# in plaintext on disk or embedded in source code.
# ---------------------------------------------------------------------------
function Get-FSApiKey {
    try {
        Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
        $targetName = "HDCompanion_FreshserviceAPIKey"
        # Use CredRead via P/Invoke through a small inline C# type
        $credType = Add-Type -MemberDefinition @'
[DllImport("advapi32.dll", EntryPoint="CredReadW", CharSet=CharSet.Unicode, SetLastError=true)]
public static extern bool CredRead(string target, int type, int flags, out IntPtr credential);

[DllImport("advapi32.dll", EntryPoint="CredFree")]
public static extern void CredFree(IntPtr credential);

[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct CREDENTIAL {
    public uint Flags;
    public uint Type;
    public string TargetName;
    public string Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public string TargetAlias;
    public string UserName;
}
'@ -Name 'CredManager' -Namespace 'HDC' -PassThru -ErrorAction Stop

        $credPtr = [IntPtr]::Zero
        if ($credType::CredRead($targetName, 1, 0, [ref]$credPtr)) {
            $cred = [System.Runtime.InteropServices.Marshal]::PtrToStructure($credPtr, [HDC.CredManager+CREDENTIAL])
            $blobBytes = New-Object byte[] $cred.CredentialBlobSize
            [System.Runtime.InteropServices.Marshal]::Copy($cred.CredentialBlob, $blobBytes, 0, $cred.CredentialBlobSize)
            $credType::CredFree($credPtr)
            return [System.Text.Encoding]::Unicode.GetString($blobBytes)
        }
    } catch {}
    return $null
}

function Set-FSApiKey {
    param([Parameter(Mandatory=$true)][string]$ApiKey)
    try {
        $targetName = "HDCompanion_FreshserviceAPIKey"
        $credType = Add-Type -MemberDefinition @'
[DllImport("advapi32.dll", EntryPoint="CredWriteW", CharSet=CharSet.Unicode, SetLastError=true)]
public static extern bool CredWrite(ref CREDENTIAL credential, uint flags);

[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct CREDENTIAL {
    public uint Flags;
    public uint Type;
    public string TargetName;
    public string Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public string TargetAlias;
    public string UserName;
}
'@ -Name 'CredManagerWrite' -Namespace 'HDC' -PassThru -ErrorAction Stop

        $blobBytes = [System.Text.Encoding]::Unicode.GetBytes($ApiKey)
        $blobPtr   = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($blobBytes.Length)
        [System.Runtime.InteropServices.Marshal]::Copy($blobBytes, 0, $blobPtr, $blobBytes.Length)

        $cred = New-Object HDC.CredManagerWrite+CREDENTIAL
        $cred.Type               = 1   # CRED_TYPE_GENERIC
        $cred.TargetName         = $targetName
        $cred.Comment            = "HelpDesk Companion Freshservice API Key"
        $cred.CredentialBlobSize = $blobBytes.Length
        $cred.CredentialBlob     = $blobPtr
        $cred.Persist            = 2   # CRED_PERSIST_LOCAL_MACHINE

        $result = $credType::CredWrite([ref]$cred, 0)
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($blobPtr)
        return $result
    } catch {
        Write-Warning "Failed to save API key to Credential Manager: $($_.Exception.Message)"
        return $false
    }
}

function Get-AppConfig {
    $centralConfigPath = "\\vm-isserver\toolkit\[it toolkit]\hdcompanion\hdcompanioncfg.json"
    $localConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) "hdcompanioncfg.json"
    
    $finalConfig = $null

    # 1. Attempt to load Central Network Config
    $isNetworkAvailable = $false
    if ($centralConfigPath -match "^\\\\([^\\]+)") {
        $serverName = $matches[1]
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            if ($ping.Send($serverName, 500).Status -eq "Success") { $isNetworkAvailable = $true }
        } catch {}
    }

    if ($isNetworkAvailable -and (Test-Path -LiteralPath $centralConfigPath)) {
        try {
            $finalConfig = Get-Content -LiteralPath $centralConfigPath -Raw | ConvertFrom-Json
            $finalConfig | Add-Member -MemberType NoteProperty -Name "LoadedConfigPath" -Value $centralConfigPath -Force
            $Script:CurrentConfigPath = $centralConfigPath
        } catch {}
    }

    # 2. Attempt to load Local Config (App Root Directory)
    if ($null -eq $finalConfig -and (Test-Path -LiteralPath $localConfigPath)) {
        try {
            $finalConfig = Get-Content -LiteralPath $localConfigPath -Raw | ConvertFrom-Json
            $finalConfig | Add-Member -MemberType NoteProperty -Name "LoadedConfigPath" -Value $localConfigPath -Force
            $Script:CurrentConfigPath = $localConfigPath
        } catch {}
    } 

    # 3. Fallback to Embedded Defaults (Slate Theme)
    if ($null -eq $finalConfig) {
        $Script:CurrentConfigPath = "Embedded Defaults (Slate Theme)"
        $embeddedJson = @'
{
    "GeneralSettings": { "DomainName": "pscu.local", "LogDirectoryUNC": "\\\\vm-isserver\\Toolkit\\[IT Toolkit]\\HDCompanion\\Logs", "LogRetentionDays": 90, "FilteredUsers": ["guest", "support", "admbrian"], "DefaultTheme": "Light", "SplashtopAPIToken": "", "FreshserviceDomain": "https://pelicanstatecreditunion.freshservice.com", "FreshserviceAPIKey": "" },
    "AutoSettings": { "AutoRefreshOptions": ["30 seconds", "1 minute", "2 minutes", "5 minutes", "10 minutes", "15 minutes", "30 minutes"], "AutoUnlockIntervalSeconds": 300 },
    "EmailSettings": { "EnableEmailNotifications": true, "SmtpServer": "10.104.100.165", "SmtpPort": 25, "EnableSsl": true, "FromAddress": "AccountMonitor@pelicancu.com", "ToAddress": ["tdawsey@pelicancu.com", "bhigginbotham@pelicancu.com", "adaigle@pelicancu.com"], "SmtpUsername": "", "SmtpPassword": "" },
    "LightModeColors": { "Text": [15, 23, 42], "Primary": [37, 99, 235], "Danger": [220, 38, 38], "TextSecondary": [100, 116, 139], "Background": [248, 250, 252], "Success": [5, 150, 105], "Card": [255, 255, 255], "Secondary": [203, 213, 225], "Hover": [241, 245, 249], "OnlineText": [5, 150, 105], "OfflineText": [220, 38, 38] },
    "DarkModeColors": { "Text": [248, 250, 252], "Primary": [59, 130, 246], "Danger": [239, 68, 68], "TextSecondary": [148, 163, 184], "Background": [11, 17, 32], "Success": [16, 185, 129], "Card": [30, 41, 59], "Secondary": [51, 65, 85], "Hover": [39, 53, 76], "OnlineText": [16, 185, 129], "OfflineText": [239, 68, 68], "OnlineBg": [6, 78, 59], "OfflineBg": [127, 29, 29] },
    "ControlProperties": { "TitleLabel": { "Text": "HelpDesk Companion" }, "SubtitleLabel": { "Text": "Less searching. More solving." }, "UnlockButton": { "Text": "Unlock Selected" }, "UnlockAllButton": { "Text": "Unlock All" }, "RefreshButton": { "Text": "Refresh" }, "SearchButton": { "Text": "Search" }, "ViewLogButton": { "Text": "View Logs" } },
    "LoadedConfigPath": "Embedded Defaults"
}
'@
        $finalConfig = ($embeddedJson | ConvertFrom-Json)
    }

    # 4. Apply User Preferences (Master + Local Override)
    $prefsFile = Join-Path $env:LOCALAPPDATA "PelicanCU\HDCompanion\userprefs.json"
    
    # Initialize the container for local prefs
    $finalConfig | Add-Member -MemberType NoteProperty -Name "UserPreferences" -Value @{
        Density = "Comfortable"
        FontSize = "Default"
        GlassEffect = $false
    } -Force

    if (Test-Path -LiteralPath $prefsFile) {
        try {
            $prefs = Get-Content -LiteralPath $prefsFile -Raw | ConvertFrom-Json
            if ($null -ne $prefs.DefaultTheme) { $finalConfig.GeneralSettings.DefaultTheme = $prefs.DefaultTheme }
            if ($null -ne $prefs.LightModeColors -and $null -ne $prefs.LightModeColors.Primary) { 
                $finalConfig.LightModeColors.Primary = $prefs.LightModeColors.Primary 
            }
            if ($null -ne $prefs.DarkModeColors -and $null -ne $prefs.DarkModeColors.Primary) { 
                $finalConfig.DarkModeColors.Primary = $prefs.DarkModeColors.Primary 
            }
            
            # Map new UI options
            if ($null -ne $prefs.Density) { $finalConfig.UserPreferences.Density = $prefs.Density }
            if ($null -ne $prefs.FontSize) { $finalConfig.UserPreferences.FontSize = $prefs.FontSize }
            if ($null -ne $prefs.GlassEffect) { $finalConfig.UserPreferences.GlassEffect = $prefs.GlassEffect }
            
            $finalConfig.LoadedConfigPath += " (+ UserPrefs)"
        } catch {
            Write-Warning "Failed to load user preferences."
        }
    }

    # 5. Overlay Freshservice API key from Windows Credential Manager
    # The key is NEVER stored in JSON files. It lives exclusively in Credential Manager
    # under the target name "HDCompanion_FreshserviceAPIKey".
    # If no key is found in Credential Manager, fall back to the value that may have
    # been set in the network config file (allowing a one-time migration path).
    $credKey = Get-FSApiKey
    if (-not [string]::IsNullOrWhiteSpace($credKey)) {
        $finalConfig.GeneralSettings.FreshserviceAPIKey = $credKey
    } elseif ([string]::IsNullOrWhiteSpace($finalConfig.GeneralSettings.FreshserviceAPIKey)) {
        Write-Warning "HDCompanion: No Freshservice API key found in Credential Manager. Freshservice features will be unavailable. Run Set-FSApiKey to configure."
    }

    return $finalConfig
}

function Initialize-LogDirectory {
    param($Config)
    $path = $Config.GeneralSettings.LogDirectoryUNC
    
    if (-not (Test-Path -LiteralPath $path)) {
        try {
            $escapedPath = $path.Replace("[", "`[").Replace("]", "`]")
            New-Item -Path $escapedPath -ItemType Directory -Force | Out-Null
        } catch {
            $fallback = Join-Path $env:TEMP "ADUnlockLogs"
            $Config.GeneralSettings.LogDirectoryUNC = $fallback
            if (-not (Test-Path -LiteralPath $fallback)) {
                New-Item -Path $fallback -ItemType Directory -Force | Out-Null
            }
        }
    }
}

function Add-AppLog {
    param($Event, $Username, $Details, $Config, $State, $Status="Success", $Color="Black")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $fileTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
    $msg = "[$timestamp] $Event - ${Username}: $Details"
    
    # Resolve txtLog safely -- UIControls may be null if called before the UI is
    # fully initialised (e.g. from a background job tick or a module-level log call).
    $txtLog = $null
    $uiControls = if ($null -ne $State) { $State["UIControls"] } else { $null }
    if ($null -ne $uiControls) { $txtLog = $uiControls["txtLog"] }

    if ($txtLog) {
        $txtLog.Dispatcher.Invoke({
            $paragraph = New-Object System.Windows.Documents.Paragraph
            $paragraph.Margin = "0"
            $run = New-Object System.Windows.Documents.Run($msg)
            
            if ($State.CurrentTheme -eq "Dark") {
                if ($Color -eq "Black") {  $run.Foreground = [System.Windows.Media.Brushes]::White  } 
                else {
                    switch ($Color) {
                        "Blue"   { $run.Foreground = [System.Windows.Media.Brushes]::DeepSkyBlue }
                        "Red"    { $run.Foreground = [System.Windows.Media.Brushes]::LightCoral }
                        "Green"  { $run.Foreground = [System.Windows.Media.Brushes]::LightGreen }
                        "Orange" { $run.Foreground = [System.Windows.Media.Brushes]::Orange }
                        Default { $run.Foreground = [System.Windows.Media.Brushes]::White }
                    }
                }
            } else {
                if ($Color -eq "Black") {  $run.Foreground = [System.Windows.Media.Brushes]::Black  } 
                else {
                    try {
                        $brushConverter = New-Object System.Windows.Media.BrushConverter
                        $run.Foreground = $brushConverter.ConvertFromString($Color)
                        if ($Color -eq "Green" -or $Color -eq "Red" -or $Color -eq "Blue") { $run.FontWeight = "Bold" }
                    } catch { $run.Foreground = [System.Windows.Media.Brushes]::Black }
                }
            }

            $paragraph.Inlines.Add($run)
            $txtLog.Document.Blocks.Add($paragraph)
            $txtLog.ScrollToEnd()
        })
    }

    try {
        $logDir = $Config.GeneralSettings.LogDirectoryUNC
        $today = Get-Date -Format "yyyyMMdd"
        $logFile = Join-Path $logDir "UnlockLog_$today.csv"
        $operator = $env:USERNAME
        $machine = $env:COMPUTERNAME

        $obj = [PSCustomObject]@{ Timestamp = $fileTimestamp; Event = $Event; Username = $Username; Details = $Details; Status = $Status; Operator = $operator; MachineName = $machine }
        $obj | Export-Csv -LiteralPath $logFile -NoTypeInformation -Append -Encoding UTF8
    } catch {}
}

function Get-AppLogFiles {
    param($Config)
    $logDir = $Config.GeneralSettings.LogDirectoryUNC
    $allLogs = @()
    if (Test-Path -LiteralPath $logDir) {
        $files = Get-ChildItem -LiteralPath $logDir -Filter "UnlockLog_*.csv"
        foreach ($f in $files) {
            $content = Import-Csv -LiteralPath $f.FullName
            $allLogs += $content
        }
    }
    return $allLogs
}

function Get-FSAssetDetails {
    param($AssetName, $Config)
    $domain = $Config.GeneralSettings.FreshserviceDomain
    if (-not $domain) { $domain = "https://pelicanstatecreditunion.freshservice.com" }
    $domain = $domain.TrimEnd('/')
    if (-not $domain.StartsWith("http")) { $domain = "https://$domain" }
    
    $apiKey = $Config.GeneralSettings.FreshserviceAPIKey
    if (-not $apiKey) { throw "Freshservice API Key is missing. Please add 'FreshserviceAPIKey' to your config." }
    
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
    $headers = @{ "Authorization" = "Basic $encoded" }
    $AssetQS = "name:'$AssetName'"
    $AssetEncoded = [Uri]::EscapeDataString("`"$AssetQS`"")
    $AssetURL = "$domain/api/v2/assets?query=$AssetEncoded"
    
    try {
        $AssetResp = Invoke-RestMethod -Uri $AssetURL -Headers $headers -Method Get -ErrorAction Stop
        if ($AssetResp.assets.Count -gt 0) { return ($AssetResp.assets | Sort-Object { [DateTime]$_.updated_at } -Descending | Select-Object -First 1) }
        return $null
    } catch {
        if ($_.Exception.Response.StatusCode -eq 'NotFound' -or $_.Exception.Message -match "404") { throw "HTTP 404 Not Found. Ensure your domain is correct in settings." } 
        else { throw $_.Exception.Message }
    }
}

function Get-FSUserAsset {
    param($User, $Config)
    $domain = $Config.GeneralSettings.FreshserviceDomain
    if (-not $domain) { $domain = "https://pelicanstatecreditunion.freshservice.com" }
    $domain = $domain.TrimEnd('/')
    if (-not $domain.StartsWith("http")) { $domain = "https://$domain" }
    
    $apiKey = $Config.GeneralSettings.FreshserviceAPIKey
    if (-not $apiKey) { throw "Freshservice API Key is missing. Please add 'FreshserviceAPIKey' to your config." }
    
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
    $headers = @{ "Authorization" = "Basic $encoded" }

    try {
        if ([string]::IsNullOrWhiteSpace($User.EmailAddress)) {
            throw "User does not have an Email Address in Active Directory. Freshservice lookup requires an email."
        }

        # 1. Look up Requester by Email Address using the correct V2 query format
        $reqId = $null
        $email = $User.EmailAddress.Trim()
        $reqQuery = "primary_email:'$email'"
        $reqEncoded = [Uri]::EscapeDataString("`"$reqQuery`"")
        $ReqURL = "$domain/api/v2/requesters?query=$reqEncoded"

        $ReqResp = Invoke-RestMethod -Uri $ReqURL -Headers $headers -Method Get -ErrorAction Stop

        if ($ReqResp.requesters -and $ReqResp.requesters.Count -gt 0) { 
            $reqId = $ReqResp.requesters[0].id 
        } else {
            throw "User email '$email' was not found in the Freshservice Requesters database."
        }

        # 2. Query assets assigned to this specific user ID
        if ($reqId) {
            $AssetQS = "user_id:$reqId"
            $AssetEncoded = [Uri]::EscapeDataString("`"$AssetQS`"")
            $AssetURL = "$domain/api/v2/assets?query=$AssetEncoded"
            $AssetResp = Invoke-RestMethod -Uri $AssetURL -Headers $headers -Method Get -ErrorAction Stop
            
            if ($AssetResp.assets -and $AssetResp.assets.Count -gt 0) { 
                return $AssetResp.assets 
            }
        }
        
        return @()
    } catch {
        $errMsg = $_.Exception.Message
        
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $respBody = $reader.ReadToEnd()
                $errMsg += "`n`nAPI Response:`n$respBody"
            } catch {}
        }

        if ($_.Exception.Response.StatusCode -eq 'NotFound' -or $errMsg -match "404") { 
            throw "HTTP 404 Not Found. Ensure your domain is correct in settings." 
        } else { 
            throw $errMsg 
        }
    }
}

function Get-FSUserRecord {
    param($User, $Config)
    
    $domain = $Config.GeneralSettings.FreshserviceDomain
    if (-not $domain) { $domain = "https://pelicanstatecreditunion.freshservice.com" }
    $domain = $domain.TrimEnd('/')
    if (-not $domain.StartsWith("http")) { $domain = "https://$domain" }
    
    $apiKey = $Config.GeneralSettings.FreshserviceAPIKey
    if (-not $apiKey) { throw "Freshservice API Key is missing. Please add 'FreshserviceAPIKey' to your config." }
    
    if ([string]::IsNullOrWhiteSpace($User.EmailAddress)) {
        throw "User does not have an Email Address in Active Directory. Freshservice lookup requires an email."
    }

    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
    $headers = @{ "Authorization" = "Basic $encoded" }
    $email = $User.EmailAddress.Trim()

    # 1. Try finding them as a Requester
    try {
        $reqQuery = "primary_email:'$email'"
        $reqEncoded = [Uri]::EscapeDataString("`"$reqQuery`"")
        $ReqURL = "$domain/api/v2/requesters?query=$reqEncoded"
        
        $ReqResp = Invoke-RestMethod -Uri $ReqURL -Headers $headers -Method Get -ErrorAction Stop
        
        if ($ReqResp.requesters -and $ReqResp.requesters.Count -gt 0) {
            $reqId = $ReqResp.requesters[0].id
            return [PSCustomObject]@{
                Type = "Requester"
                Id = $reqId
                Url = "$domain/itil/requesters/$reqId"
            }
        }
    } catch {
        # Suppress errors on the first pass, we will try the Agent database next
    }

    # 2. Try finding them as an Agent (If they aren't a standard requester)
    try {
        $agentEncoded = [Uri]::EscapeDataString($email)
        $AgentURL = "$domain/api/v2/agents?email=$agentEncoded"
        
        $AgentResp = Invoke-RestMethod -Uri $AgentURL -Headers $headers -Method Get -ErrorAction Stop
        
        if ($AgentResp.agents -and $AgentResp.agents.Count -gt 0) {
            $agentId = $AgentResp.agents[0].id
            return [PSCustomObject]@{
                Type = "Agent"
                Id = $agentId
                Url = "$domain/admin/agents/$agentId/edit"
            }
        }
    } catch {
        throw "Error querying the Agents database: $($_.Exception.Message)"
    }

    return $null
}

function Get-FluentThemeColors {
    param($State)
    $isDark = ($null -ne $State -and $State.CurrentTheme -eq "Dark")
    
    # Defaults (Slate)
    $primaryHexDark = "#3B82F6"
    $primaryHexLight = "#2563EB"

    # Dynamically inject custom User Pref Hex Colors if active
    if ($global:Config) {
        try {
            $pLight = $global:Config.LightModeColors.Primary
            if ($pLight -is [string]) { $primaryHexLight = $pLight }
            elseif ($pLight -and $pLight.Count -eq 3) { $primaryHexLight = "#{0:X2}{1:X2}{2:X2}" -f $pLight[0], $pLight[1], $pLight[2] }
            
            $pDark = $global:Config.DarkModeColors.Primary
            if ($pDark -is [string]) { $primaryHexDark = $pDark }
            elseif ($pDark -and $pDark.Count -eq 3) { $primaryHexDark = "#{0:X2}{1:X2}{2:X2}" -f $pDark[0], $pDark[1], $pDark[2] }
        } catch {}
    }

    return @{
        Bg         = if ($isDark) { "#0B1120" } else { "#F8FAFC" }
        Fg         = if ($isDark) { "#F8FAFC" } else { "#0F172A" }
        SecFg      = if ($isDark) { "#94A3B8" } else { "#64748B" }
        BtnBg      = if ($isDark) { "#1E293B" } else { "#FFFFFF" }
        BtnBorder  = if ($isDark) { "#334155" } else { "#CBD5E1" }
        PrimaryBg  = if ($isDark) { $primaryHexDark } else { $primaryHexLight }
        PrimaryFg  = if ($isDark) { "#FFFFFF" } else { "#FFFFFF" }
        GridBorder = if ($isDark) { "#1E293B" } else { "#E2E8F0" }
        HoverBg    = if ($isDark) { "#27354C" } else { "#F1F5F9" }
        AltRowBg   = if ($isDark) { "#1E293B" } else { "#FFFFFF" }
        Danger     = if ($isDark) { "#EF4444" } else { "#DC2626" }
    }
}

function Load-XamlWindow {
    param(
        [Parameter(Mandatory=$true)]
        [string]$XamlPath,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$ThemeColors
    )
    
    if (-not (Test-Path -LiteralPath $XamlPath)) {
        throw "XAML file not found at path: $XamlPath"
    }

    $xamlText = [System.IO.File]::ReadAllText($XamlPath)
    
    # Inject Theme Colors directly into the XAML markup before parsing
    if ($ThemeColors) {
        foreach ($key in $ThemeColors.Keys) {
            $xamlText = $xamlText -replace "\{Theme_$key\}", $ThemeColors[$key]
        }
    }

    # Custom DTD Processing override required for modern .NET framework security updates
    $xmlSettings = New-Object System.Xml.XmlReaderSettings
    $xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Parse
    
    $stringReader = [System.IO.StringReader]::new($xamlText)
    $xmlReader = [System.Xml.XmlReader]::Create($stringReader, $xmlSettings)
    
    try {
        return [System.Windows.Markup.XamlReader]::Load($xmlReader)
    } finally {
        $xmlReader.Close()
        $stringReader.Close()
    }
}

Export-ModuleMember -Function Get-AppConfig, Initialize-LogDirectory, Add-AppLog, Get-AppLogFiles, Get-FSAssetDetails, Get-FSUserAsset, Get-FSUserRecord, Get-FluentThemeColors, Load-XamlWindow, Get-FSApiKey, Set-FSApiKey