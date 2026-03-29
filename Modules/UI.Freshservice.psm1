# ============================================================================
# UI_Freshservice.psm1 - Freshservice Integration Interface Logic
# ============================================================================

# ---------------------------------------------------------------------------
# Internal helpers -- resolve Freshservice API credentials from the config
# object. The single authoritative location is GeneralSettings, where
# CoreLogic.psm1 overlays the key from Windows Credential Manager at startup.
# These helpers eliminate the 6-path waterfall that previously lived inside
# Submit-FSQuickTicket and keep all credential resolution in one place.
# ---------------------------------------------------------------------------
function Get-FSApiToken {
    param([Parameter(Mandatory=$true)]$Config)
    # Primary path: GeneralSettings (populated by CoreLogic from Credential Manager)
    if ($Config.GeneralSettings -and -not [string]::IsNullOrWhiteSpace($Config.GeneralSettings.FreshserviceAPIKey)) {
        return $Config.GeneralSettings.FreshserviceAPIKey
    }
    throw "Freshservice API key is missing. Run Set-FSApiKey to store it in Windows Credential Manager."
}

function Get-FSApiUrl {
    param([Parameter(Mandatory=$true)]$Config)
    if ($Config.GeneralSettings -and -not [string]::IsNullOrWhiteSpace($Config.GeneralSettings.FreshserviceDomain)) {
        $url = $Config.GeneralSettings.FreshserviceDomain.TrimEnd('/')
        if (-not $url.StartsWith("http")) { $url = "https://$url" }
        return $url
    }
    throw "Freshservice domain URL is missing from GeneralSettings.FreshserviceDomain in config."
}

function Submit-FSQuickTicket {
    param($RequesterEmail, $Subject, $Description, $Status, $Config, $Category, $SubCategory, $ItemCategory, $CustomFields, $AssigneeEmail)

    $statusMap = @{ "Open"=2; "Pending"=3; "Resolved"=4; "Closed"=5 }
    $statusId  = if ($statusMap.Contains($Status)) { $statusMap[$Status] } else { 4 }

    $token = Get-FSApiToken -Config $Config
    $url   = Get-FSApiUrl   -Config $Config

    $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($token):X"))
    $headers = @{ "Authorization" = "Basic $encodedCreds"; "Content-Type" = "application/json" }
    
    $responderId = $null
    if (-not [string]::IsNullOrWhiteSpace($AssigneeEmail) -and $AssigneeEmail -ne "Unassigned") {
        try {
            $agentRes = Invoke-RestMethod -Uri "$url/api/v2/agents?email=$AssigneeEmail" -Method Get -Headers $headers -ErrorAction Stop
            if ($agentRes.agents -and $agentRes.agents.Count -gt 0) { $responderId = $agentRes.agents[0].id }
        } catch { Write-Warning "Could not resolve Assignee email to Agent ID." }
    }

    $bodyHash = @{
        description    = $Description
        subject        = $Subject
        status         = $statusId
        priority       = 1
        source         = 2
        email          = $RequesterEmail
        category       = $Category
        sub_category   = $SubCategory
        item_category  = $ItemCategory
        custom_fields  = $CustomFields
    }
    if ($responderId) { $bodyHash.Add("responder_id", [long]$responderId) }
    $body = $bodyHash | ConvertTo-Json -Depth 5
    
    try {
        $res = Invoke-RestMethod -Uri "$url/api/v2/tickets" -Method Post -Headers $headers -Body $body -ErrorAction Stop
        return [string]$res.ticket.id
    } catch {
        $errorMsg = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $rawResponse = $reader.ReadToEnd()
                $json = $rawResponse | ConvertFrom-Json
                if ($json.description) { $errorMsg += "`n`nFS Reason: $($json.description)" }
                if ($json.errors) { foreach ($err in $json.errors) { $errorMsg += "`n- $($err.field): $($err.message)" } }
            } catch { }
        }
        throw $errorMsg
    }
}

function Register-FreshserviceUIEvents {
    param($Window, $Config, $State)

    $lvData = $Window.FindName("lvData")
    $ctxFSInventory = $Window.FindName("ctxFSInventory")
    $ctxFindComputer = $Window.FindName("ctxFindComputer")
    $ctxOpenFSRecord = $Window.FindName("ctxOpenFSRecord")
    $ctxTicketFeed   = $Window.FindName("ctxTicketFeed")
    $ctxQuickTicket  = $Window.FindName("ctxQuickTicket")
    $ctxSep1         = $Window.FindName("ctxSep1")
    $ctxSep2         = $Window.FindName("ctxSep2")

    if ($lvData -and $lvData.ContextMenu) {
        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            $isUser = ($sel -and $sel.Type -eq "User")
            $isComp = ($sel -and $sel.Type -eq "Computer")

            if ($ctxQuickTicket)  { $ctxQuickTicket.Visibility  = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxOpenFSRecord) { $ctxOpenFSRecord.Visibility  = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxFindComputer) { $ctxFindComputer.Visibility  = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxTicketFeed)   { $ctxTicketFeed.Visibility    = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxFSInventory)  { $ctxFSInventory.Visibility   = if ($isComp) { "Visible" } else { "Collapsed" } }
            # Show separators only when both groups they divide are visible
            if ($ctxSep1) { $ctxSep1.Visibility = if ($sel)     { "Visible" } else { "Collapsed" } }
            if ($ctxSep2) { $ctxSep2.Visibility = if ($isUser)  { "Visible" } else { "Collapsed" } }
        }.GetNewClosure())
    }

    if ($ctxQuickTicket) {
        $ctxQuickTicket.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $userObj = $lvData.SelectedItem; $colors = Get-FluentThemeColors $State
                $reqEmail = if ($userObj.Email) { $userObj.Email } else { "$($userObj.Name)@pelicancu.com" }
                
                $tktXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Create Quick Ticket" Width="550" Height="700" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Window.Resources>
                        <Style TargetType="TextBox"><Setter Property="Background" Value="{Theme_BtnBg}"/><Setter Property="Foreground" Value="{Theme_Fg}"/><Setter Property="BorderBrush" Value="{Theme_BtnBorder}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="6,4"/></Style>
                        <Style TargetType="ComboBox"><Setter Property="Background" Value="{Theme_BtnBg}"/><Setter Property="Foreground" Value="{Theme_Fg}"/><Setter Property="BorderBrush" Value="{Theme_BtnBorder}"/><Setter Property="BorderThickness" Value="1"/></Style>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value>
                            </Setter>
                        </Style>
                    </Window.Resources>
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                        <Grid>
                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                            <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand"><TextBlock Text="Log Quick Ticket" FontSize="16" FontWeight="SemiBold" Foreground="{Theme_Fg}"/></Border>
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="16,8,16,16">
                                    <TextBlock Text="Requester Email:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <TextBox Text="$reqEmail" IsReadOnly="True" Margin="0,0,0,12" Background="Transparent" BorderThickness="0" FontWeight="Bold"/>
                                    <TextBlock Text="Subject:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtSubject" Text="" Margin="0,0,0,12"/>
                                    <TextBlock Text="Notes / Description:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtDesc" Text="" AcceptsReturn="True" TextWrapping="Wrap" Height="60" Margin="0,0,0,12"/>
                                    <TextBlock Text="Ticket Status:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <ComboBox x:Name="cbStatus" Height="30" SelectedIndex="0" Margin="0,0,0,16"><ComboBoxItem Content="Open"/><ComboBoxItem Content="Pending"/><ComboBoxItem Content="Resolved"/><ComboBoxItem Content="Closed"/></ComboBox>
                                    
                                    <TextBlock Text="Required Ticket Properties:" FontWeight="Bold" Foreground="{Theme_Fg}" Margin="0,0,0,8"/>
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="15"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                        
                                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Category:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbCategory" Grid.Row="1" Grid.Column="0" Height="28" Margin="0,0,0,12"/>
                                        <TextBlock Grid.Row="0" Grid.Column="2" Text="Sub-Category:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="txtSubCategory" Grid.Row="1" Grid.Column="2" Height="28" IsEditable="True" Text="" Margin="0,0,0,12"/>
                                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Item:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="txtItemCategory" Grid.Row="3" Grid.Column="0" Height="28" IsEditable="True" Text="" Margin="0,0,0,12"/>
                                        <TextBlock Grid.Row="2" Grid.Column="2" Text="Resolved Remotely:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbResolvedRemotely" Grid.Row="3" Grid.Column="2" Height="28" Margin="0,0,0,12"><ComboBoxItem Content="Yes" IsSelected="True"/><ComboBoxItem Content="No"/></ComboBox>
                                        <TextBlock Grid.Row="4" Grid.Column="0" Text="Member Impacting:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbMemberImpacting" Grid.Row="5" Grid.Column="0" Height="28" Margin="0,0,0,12"><ComboBoxItem Content="Yes"/><ComboBoxItem Content="No" IsSelected="True"/></ComboBox>
                                        <TextBlock Grid.Row="4" Grid.Column="2" Text="Who is Affected:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbWhoAffected" Grid.Row="5" Grid.Column="2" Height="28" Margin="0,0,0,12"><ComboBoxItem Content="You Only" IsSelected="True"/><ComboBoxItem Content="Entire Department/Branch"/><ComboBoxItem Content="Company Wide"/></ComboBox>
                                        <TextBlock Grid.Row="6" Grid.Column="0" Text="Prevents Crit. Operations:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbPreventsOps" Grid.Row="7" Grid.Column="0" Height="28" Margin="0,0,0,4"><ComboBoxItem Content="Yes"/><ComboBoxItem Content="No" IsSelected="True"/></ComboBox>
                                        <TextBlock Grid.Row="6" Grid.Column="2" Text="Is there a Workaround:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbWorkaround" Grid.Row="7" Grid.Column="2" Height="28" Margin="0,0,0,4"><ComboBoxItem Content="Yes" IsSelected="True"/><ComboBoxItem Content="No"/></ComboBox>
                                        <TextBlock Grid.Row="8" Grid.Column="0" Text="Assignee:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbAssignee" Grid.Row="9" Grid.Column="0" Height="28" Margin="0,0,0,4" IsEditable="True"/>
                                    </Grid>
                                </StackPanel>
                            </ScrollViewer>
                            <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}">
                                <Grid>
                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                        <Button x:Name="btnSubTkt" Content="Submit Ticket" Width="100" Height="28" Margin="0,0,8,0" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsDefault="True"/>
                                        <Button x:Name="btnCancelTkt" Content="Cancel" Width="80" Height="28" Background="{Theme_Bg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" IsCancel="True"/>
                                    </StackPanel>
                                    <Thumb x:Name="thumbResizeTkt" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="12" Height="12" Cursor="SizeNWSE" Margin="0,0,-8,-8" Background="Transparent" ToolTip="Resize Window"/>
                                </Grid>
                            </Border>
                        </Grid>
                    </Border>
                </Window>
"@
                $xamlText = $tktXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $tktWin = [System.Windows.Markup.XamlReader]::Load($reader)
                $tktWin.Owner = $Window
                
                $tktWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $tktWin.DragMove() }.GetNewClosure())
                $thumbResizeTkt = $tktWin.FindName("thumbResizeTkt")
                if ($thumbResizeTkt) {
                    $thumbResizeTkt.Add_DragDelta({ param($sender, $e)
                        $newWidth = $tktWin.Width + $e.HorizontalChange; $newHeight = $tktWin.Height + $e.VerticalChange
                        if ($newWidth -gt 400) { $tktWin.Width = $newWidth }
                        if ($newHeight -gt 350) { $tktWin.Height = $newHeight }
                    }.GetNewClosure())
                }

                $txtSub = $tktWin.FindName("txtSubject"); $txtDesc = $tktWin.FindName("txtDesc"); $cbStatus = $tktWin.FindName("cbStatus")
                $txtSubCat = $tktWin.FindName("txtSubCategory"); $txtItemCat = $tktWin.FindName("txtItemCategory")
                $cbCategory = $tktWin.FindName("cbCategory")
                $cats = @('Access & Security','Accounting','Account Servicing','Alerts','Auditing','Building and Grounds Maintenance','Change Request','Card Services','Cyber Security Incident','Development & Reporting','Digital Banking','Documents','Email','Facilities','File & Folder','Genesys Cloud Change Management','Hardware','Human Resources','Microsoft Authenticator','Mobile Device Management','Morning Tasks','Network','Office Furniture','Purchasing','Quality Assurance (QA)','Relocations','Scheduled Maintenance','Software','Time Tracking','Training','User Accounts','VDI','Video Playback','WFH Application','Freshservice','Member Experience','Project Management','ITM/ATM')
                if ($cbCategory) { $cbCategory.ItemsSource = $cats; $cbCategory.SelectedItem = 'User Accounts' }

                # Category -> SubCategory -> Item lookup table
                # SubCategory keys map to arrays of Items
                $catMap = @{
                    'Access & Security'         = [ordered]@{ 'Multi-Factor Authentication'=@('DUO Setup','Microsoft Authenticator','RSA Token'); 'Permissions'=@('Share Access','Folder Access','Application Access','Remove Access'); 'Password'=@('Password Reset','Account Unlock','Password Complexity'); 'VPN'=@('Setup','Connectivity','Split Tunnel') }
                    'Accounting'                = [ordered]@{ 'Accounts Payable'=@('Invoice Processing','Vendor Payment','Expense Report'); 'Accounts Receivable'=@('Payment Application','Collections','Statements'); 'General Ledger'=@('Journal Entry','Reconciliation','Reporting') }
                    'Account Servicing'         = [ordered]@{ 'Loan Servicing'=@('Payment Processing','Payoff Request','Escrow'); 'Deposit Accounts'=@('Account Maintenance','Statement Request','Fee Waiver'); 'Cards'=@('Debit Card','Credit Card','Dispute') }
                    'Alerts'                    = [ordered]@{ 'System Alerts'=@('Server Alert','Network Alert','Security Alert'); 'Member Alerts'=@('Fraud Alert','Overdraft Alert','Balance Alert') }
                    'Auditing'                  = [ordered]@{ 'Internal Audit'=@('Process Review','Compliance Check','Documentation'); 'External Audit'=@('Exam Prep','Document Request','Findings Response') }
                    'Building and Grounds Maintenance' = [ordered]@{ 'HVAC'=@('Temperature Issue','Maintenance Request'); 'Plumbing'=@('Repair','Inspection'); 'Electrical'=@('Lighting','Outlets','Breaker'); 'Grounds'=@('Landscaping','Parking Lot','Signage') }
                    'Change Request'            = [ordered]@{ 'Infrastructure'=@('Server Change','Network Change','Firewall Rule'); 'Application'=@('Software Update','Configuration Change','New Deployment'); 'Process'=@('Policy Change','Procedure Update') }
                    'Card Services'             = [ordered]@{ 'Debit Cards'=@('Issuance','Activation','Dispute','Lost/Stolen'); 'Credit Cards'=@('Issuance','Activation','Dispute','Lost/Stolen','Limit Change'); 'ATM'=@('Replenishment','Maintenance','Out of Service') }
                    'Cyber Security Incident'   = [ordered]@{ 'Phishing'=@('Reported Email','Clicked Link','Data Submitted'); 'Malware'=@('Virus','Ransomware','Spyware'); 'Unauthorized Access'=@('Account Compromise','Privilege Escalation','Data Exfiltration') }
                    'Development & Reporting'   = [ordered]@{ 'Report Development'=@('New Report','Modify Existing','Ad Hoc Query'); 'Application Development'=@('New Feature','Bug Fix','Enhancement'); 'Data'=@('Data Extract','Data Quality','ETL Issue') }
                    'Digital Banking'           = [ordered]@{ 'Online Banking'=@('Login Issue','Enrollment','Feature Request','Error'); 'Mobile Banking'=@('App Issue','Enrollment','Notification'); 'Bill Pay'=@('Payment Issue','Payee Setup','History') }
                    'Documents'                 = [ordered]@{ 'Forms'=@('New Form','Update Form','E-Sign Setup'); 'Policies'=@('Policy Document','Procedure','Work Instruction'); 'Storage'=@('SharePoint','Network Share','Archive') }
                    'Email'                     = [ordered]@{ 'Microsoft Exchange'=@('New Mailbox','Distribution List','Shared Mailbox','Quota'); 'Outlook'=@('Configuration','Profile Rebuild','Rules','Calendar'); 'Spam/Phishing'=@('False Positive','Block Sender','Quarantine Release') }
                    'Facilities'                = [ordered]@{ 'Office Space'=@('Desk Assignment','Cubicle Setup','Conference Room'); 'Keys & Access'=@('Key Fob','Door Access','Safe'); 'Furniture'=@('Request','Repair','Relocation') }
                    'File & Folder'             = [ordered]@{ 'Network Share'=@('Access Request','New Folder','Permissions Change'); 'SharePoint'=@('Site Access','Library','Permissions'); 'OneDrive'=@('Sync Issue','Storage','Sharing') }
                    'Genesys Cloud Change Management' = [ordered]@{ 'Routing'=@('Queue Change','Skill Assignment','IVR Flow'); 'Agents'=@('Add Agent','Remove Agent','Profile Update'); 'Reporting'=@('New Report','Dashboard','Schedule') }
                    'Hardware'                  = [ordered]@{ 'Computer'=@('New Setup','Repair','Replacement','Upgrade'); 'Peripherals'=@('Monitor','Keyboard/Mouse','Webcam','Headset'); 'Mobile Device'=@('Phone Setup','Tablet Setup','MDM Enrollment'); 'Printer'=@('Install','Replace Toner','Paper Jam','Offline') }
                    'Human Resources'           = [ordered]@{ 'Onboarding'=@('New Hire Setup','Account Creation','Equipment Request'); 'Offboarding'=@('Account Disable','Equipment Return','Data Preservation'); 'Employee Change'=@('Name Change','Department Transfer','Role Change') }
                    'Microsoft Authenticator'   = [ordered]@{ 'Setup'=@('New Device','Additional Device','QR Code'); 'Troubleshooting'=@('Not Receiving Code','App Error','Reset'); 'Migration'=@('New Phone','Restore Backup') }
                    'Mobile Device Management'  = [ordered]@{ 'Enrollment'=@('Intune Enrollment','BYOD','Corporate Device'); 'Policy'=@('Compliance Issue','Policy Push','App Deployment'); 'Troubleshooting'=@('Sync Issue','Remote Wipe','App Issue') }
                    'Morning Tasks'             = [ordered]@{ 'System Checks'=@('Server Health','Network Check','Backup Verification'); 'Branch Opening'=@('Teller System','Vault','Terminal Check') }
                    'Network'                   = [ordered]@{ 'Connectivity'=@('No Internet','Slow Connection','Wi-Fi Issue','VPN'); 'Infrastructure'=@('Switch','Router','Firewall','Cabling'); 'DNS/DHCP'=@('DNS Resolution','IP Conflict','DHCP Scope') }
                    'Office Furniture'          = [ordered]@{ 'Request'=@('New Furniture','Replacement','Ergonomic Assessment'); 'Repair'=@('Chair','Desk','Cabinet') }
                    'Purchasing'                = [ordered]@{ 'IT Equipment'=@('Laptop','Desktop','Monitor','Accessories'); 'Software'=@('New License','Renewal','Additional Seat'); 'Vendor'=@('New Vendor','PO Request','Invoice Discrepancy') }
                    'Quality Assurance (QA)'    = [ordered]@{ 'Testing'=@('UAT','Regression','Performance'); 'Documentation'=@('Test Plan','Test Case','Defect Report'); 'Compliance'=@('Audit Support','Policy Review','Evidence Collection') }
                    'Relocations'               = [ordered]@{ 'Employee Move'=@('Desk Relocation','Branch Transfer','Remote Setup'); 'Equipment Move'=@('Computer','Printer','Phone') }
                    'Scheduled Maintenance'     = [ordered]@{ 'Planned Outage'=@('Server Maintenance','Network Maintenance','Application Update'); 'Recurring'=@('Patch Tuesday','Backup Verification','Certificate Renewal') }
                    'Software'                  = [ordered]@{ 'Installation'=@('New Install','Reinstall','Update'); 'Troubleshooting'=@('Application Error','Performance','Compatibility'); 'Licensing'=@('License Activation','License Expired','Transfer') }
                    'Time Tracking'             = [ordered]@{ 'ADP'=@('Timecard Issue','Approval','Correction'); 'Scheduling'=@('Schedule Change','Shift Swap','Time Off') }
                    'Training'                  = [ordered]@{ 'New Employee'=@('System Training','Process Training','Compliance'); 'Ongoing'=@('Refresher','New Feature','Certification') }
                    'User Accounts'             = [ordered]@{ 'Active Directory'=@('Account Unlock','Password Reset','Account Creation','Account Disable','Group Membership'); 'Windows'=@('Account Unlock','Password Reset','Profile Issue','Login Issue'); 'Application'=@('Access Request','Password Reset','Role Assignment') }
                    'VDI'                       = [ordered]@{ 'Connectivity'=@('Cannot Connect','Slow Performance','Black Screen'); 'Profile'=@('Profile Reset','Application Missing','Settings'); 'Provisioning'=@('New VDI','Resize','Decommission') }
                    'Video Playback'            = [ordered]@{ 'Streaming'=@('Buffering','No Audio','Resolution'); 'Hardware'=@('Monitor','Projector','HDMI/Display'); 'Conferencing'=@('Teams','Zoom','Webex') }
                    'WFH Application'           = [ordered]@{ 'VPN'=@('Setup','Connectivity','Certificate'); 'Remote Desktop'=@('RDP Setup','Connection Issue','Performance'); 'Equipment'=@('Laptop Issue','Peripheral','ISP Issue') }
                    'Freshservice'              = [ordered]@{ 'Tickets'=@('Create Ticket','Update Ticket','Merge Ticket'); 'Reporting'=@('Report Request','Dashboard','Analytics'); 'Configuration'=@('Agent Setup','Automation','Workflow') }
                    'Member Experience'         = [ordered]@{ 'Complaint'=@('Service Issue','Wait Time','Facility'); 'Feedback'=@('Survey','Suggestion','Compliment') }
                    'Project Management'        = [ordered]@{ 'New Project'=@('Initiation','Planning','Resource Request'); 'Active Project'=@('Status Update','Issue','Change Request'); 'Closing'=@('Completion','Lessons Learned','Documentation') }
                    'ITM/ATM'                   = [ordered]@{ 'ATM'=@('Out of Service','Cash Replenishment','Receipt Printer','Card Reader'); 'ITM'=@('Video Issue','Audio Issue','Transaction Error','Offline') }
                }

                # Helper: populate SubCategory ComboBox for the selected category
                $PopulateSubCats = {
                    param($selCat)
                    if ($txtSubCat) {
                        $txtSubCat.ItemsSource = $null
                        $txtSubCat.Text = ""
                    }
                    if ($txtItemCat) {
                        $txtItemCat.ItemsSource = $null
                        $txtItemCat.Text = ""
                    }
                    if ($selCat -and $catMap.Contains($selCat)) {
                        if ($txtSubCat) { $txtSubCat.ItemsSource = $catMap[$selCat].Keys }
                    }
                }.GetNewClosure()

                # Helper: populate Item ComboBox for the selected subcategory
                $PopulateItems = {
                    param($selCat, $selSub)
                    if ($txtItemCat) {
                        $txtItemCat.ItemsSource = $null
                        $txtItemCat.Text = ""
                    }
                    if ($selCat -and $selSub -and $catMap.Contains($selCat) -and $catMap[$selCat].Contains($selSub)) {
                        if ($txtItemCat) { $txtItemCat.ItemsSource = $catMap[$selCat][$selSub] }
                    }
                }.GetNewClosure()

                # Wire SelectionChanged on Category -> populate SubCategory
                if ($cbCategory) {
                    $cbCategory.Add_SelectionChanged({
                        $sel = if ($cbCategory.SelectedItem) { $cbCategory.SelectedItem.ToString() } else { $cbCategory.Text }
                        & $PopulateSubCats $sel
                    }.GetNewClosure())
                    # Populate subcategories for the default selection now
                    & $PopulateSubCats 'User Accounts'
                }

                # Wire SelectionChanged on SubCategory -> populate Item
                if ($txtSubCat) {
                    $txtSubCat.Add_SelectionChanged({
                        $cat = if ($cbCategory.SelectedItem) { $cbCategory.SelectedItem.ToString() } else { $cbCategory.Text }
                        $sub = if ($txtSubCat.SelectedItem) { $txtSubCat.SelectedItem.ToString() } else { $txtSubCat.Text }
                        & $PopulateItems $cat $sub
                    }.GetNewClosure())
                }

                # Pre-fill lockout-specific fields only when the account is actually locked
                $isLocked = ($userObj.LockedOut -eq $true) -or ($userObj.LockedOut -eq "True")
                if ($isLocked) {
                    if ($txtSub)     { $txtSub.Text     = "Unlock Account: $reqEmail" }
                    if ($txtDesc)    { $txtDesc.Text    = "The Microsoft account for $reqEmail has been locked due to too many bad password attempts." }
                    if ($txtSubCat)  { $txtSubCat.Text  = "Windows"; $txtSubCat.SelectedItem = "Windows" }
                    if ($txtItemCat) {
                        & $PopulateItems 'User Accounts' 'Windows'
                        $txtItemCat.Text = "Account Unlock"; $txtItemCat.SelectedItem = "Account Unlock"
                    }
                }

                $cbAssignee = $tktWin.FindName("cbAssignee")
                if ($cbAssignee) {
                    $emails = @()
                    if ($null -ne $Config.EmailSettings) {
                        foreach ($p in $Config.EmailSettings.PSObject.Properties) {
                            if ($p.Value -is [array]) { $emails += $p.Value }
                            elseif ($p.Value -is [string] -and $p.Value -match "@") { $emails += $p.Value -split ',' }
                        }
                    }
                    $emails = @("Unassigned") + ($emails | Select-Object -Unique | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "@" })
                    $cbAssignee.ItemsSource = $emails
                    $cbAssignee.SelectedIndex = 0
                }

                $txtSubCategory = $tktWin.FindName("txtSubCategory"); $txtItemCategory = $tktWin.FindName("txtItemCategory")
                $cbResolvedRemotely = $tktWin.FindName("cbResolvedRemotely"); $cbMemberImpacting = $tktWin.FindName("cbMemberImpacting")
                $cbWhoAffected = $tktWin.FindName("cbWhoAffected"); $cbPreventsOps = $tktWin.FindName("cbPreventsOps"); $cbWorkaround = $tktWin.FindName("cbWorkaround")

                $btnCancel = $tktWin.FindName("btnCancelTkt"); if ($btnCancel) { $btnCancel.Add_Click({ $tktWin.Close() }.GetNewClosure()) }
                $btnSubmit = $tktWin.FindName("btnSubTkt")
                if ($btnSubmit) {
                    $btnSubmit.Add_Click({
                        $s = $txtSub.Text; $d = $txtDesc.Text; $st = $cbStatus.Text
                        $cat = if ($cbCategory) { $cbCategory.Text } else { "" }; $subCat = if ($txtSubCategory) { $txtSubCategory.Text } else { "" }; $itemCat = if ($txtItemCategory) { $txtItemCategory.Text } else { "" }; $assignee = if ($cbAssignee) { $cbAssignee.Text } else { "Unassigned" }
                        $cFields = @{ "resolved_remotely" = $cbResolvedRemotely.Text; "is_this_directly_member_impacting" = $cbMemberImpacting.Text; "who_or_what_is_affected_by_this_impact" = $cbWhoAffected.Text; "does_this_prevent_critical_business_operations_from_continuing" = $cbPreventsOps.Text; "is_there_a_workaround_for_this_issue" = $cbWorkaround.Text }
                        $tktWin.Close()
                        
                        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                        Add-AppLog -Event "Freshservice" -Username $userObj.Name -Details "Submitting quick ticket..." -Config $Config -State $State -Status "Info"
                        try {
                            $id = Submit-FSQuickTicket -RequesterEmail $reqEmail -Subject $s -Description $d -Status $st -Config $Config -Category $cat -SubCategory $subCat -ItemCategory $itemCat -CustomFields $cFields -AssigneeEmail $assignee
                            [System.Windows.Input.Mouse]::OverrideCursor = $null
                            Show-AppMessageBox -Message "Freshservice ticket successfully created!`n`nTicket Number: $id" -Title "Ticket Created" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors
                        } catch {
                            [System.Windows.Input.Mouse]::OverrideCursor = $null
                            Show-AppMessageBox -Message "Failed to create ticket:`n`n$($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors
                        }
                    }.GetNewClosure())
                }
                Show-CenteredOnOwner -ChildWindow $tktWin -OwnerWindow $Window
                $tktWin.Show()
            }
        }.GetNewClosure())
    }

    if ($ctxOpenFSRecord) {
        $ctxOpenFSRecord.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $userObj = $lvData.SelectedItem
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                Add-AppLog -Event "Freshservice" -Username $userObj.Name -Details "Opening Freshservice record..." -Config $Config -State $State -Status "Info"
                try {
                    $record = Get-FSUserRecord -User $userObj -Config $Config
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($record) { try { Start-Process "msedge.exe" -ArgumentList @("--app=$($record.Url)") } catch { Show-AppMessageBox -Message "Failed to open Microsoft Edge." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) } }
                    else { Show-AppMessageBox -Message "No Requester or Agent record found for $($userObj.Name) in Freshservice." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                } catch {
                     [System.Windows.Input.Mouse]::OverrideCursor = $null
                     Show-AppMessageBox -Message "Error contacting Freshservice:`n`n$($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                }
            }
        }.GetNewClosure())
    }

    if ($ctxFindComputer) {
        $ctxFindComputer.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $userObj = $lvData.SelectedItem
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                Add-AppLog -Event "Freshservice" -Username $userObj.Name -Details "Querying assigned computers..." -Config $Config -State $State -Status "Info"
                try {
                    $assets = Get-FSUserAsset -User $userObj -Config $Config
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($assets -and $assets.Count -gt 0) {
                        $compList = ""; foreach ($a in $assets) { $compList += "- $($a.name)`n" }
                        Show-AppMessageBox -Message "Found $($assets.Count) computer(s) assigned to $($userObj.Name) in Freshservice:`n`n$compList" -Title "Assigned Computers" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                    } else { Show-AppMessageBox -Message "No assets currently assigned to $($userObj.Name)." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                } catch { 
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    Show-AppMessageBox -Message "Error querying Freshservice:`n`n$($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) 
                }
            }
        }.GetNewClosure())
    }

    if ($ctxFSInventory) {
        $ctxFSInventory.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $computerName = $lvData.SelectedItem.Name
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                Add-AppLog -Event "Freshservice" -Username "System" -Details "Fetching inventory for $computerName..." -Config $Config -State $State -Status "Info"
                try {
                    $asset = Get-FSAssetDetails -AssetName $computerName -Config $Config
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($asset) {
                        if ($asset.display_id) {
                            # Build CMDB URL from the configured domain rather than a hardcoded host
                            $fsBase = Get-FSApiUrl -Config $Config
                            $cmdbUrl = "$fsBase/cmdb/items/$($asset.display_id)"
                            try { Start-Process "msedge.exe" -ArgumentList @("--app=$cmdbUrl") } catch { Show-AppMessageBox -Message "Failed to open link." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                        } else { Show-AppMessageBox -Message "Asset found but missing Display ID." -Title "Warning" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                    } else { Show-AppMessageBox -Message "Computer '$computerName' not found in Freshservice inventory." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                } catch {
                     [System.Windows.Input.Mouse]::OverrideCursor = $null
                     Show-AppMessageBox -Message "Error contacting Freshservice: $($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                }
            }
        }.GetNewClosure())
    }
    # Wire ctxTicketFeed -- visible for user rows only (visibility set by the ContextMenu handler in UI_UserActions)
    $ctxTicketFeed = $Window.FindName("ctxTicketFeed")
    if ($ctxTicketFeed) {
        $ctxTicketFeed.Add_Click({
            if (-not ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User")) { return }
            $userObj = $lvData.SelectedItem
            $colors  = Get-FluentThemeColors $State

            # Validate FS is configured before showing anything
            $token = $null; $url = $null
            try { $token = Get-FSApiToken -Config $Config; $url = Get-FSApiUrl -Config $Config } catch {
                Show-AppMessageBox -Message "Freshservice is not configured.`n$($_.Exception.Message)" -Title "Not Configured" -IconType "Warning" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                return
            }

            if ([string]::IsNullOrWhiteSpace($userObj.EmailAddress)) {
                Show-AppMessageBox -Message "This user has no Email Address in AD. Freshservice lookup requires an email." -Title "No Email" -IconType "Warning" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                return
            }

            # Show loading spinner
            $loadXaml = @"
            <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                    Title="Tickets" Width="340" Height="110" WindowStartupLocation="CenterOwner"
                    WindowStyle="None" AllowsTransparency="True" Background="Transparent"
                    FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                    <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                        <TextBlock Text="Loading Freshservice tickets..." FontSize="13" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                        <ProgressBar IsIndeterminate="True" Width="260" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                    </StackPanel>
                </Border>
            </Window>
"@
            $lx = $loadXaml; foreach ($k in $colors.Keys) { $lx = $lx.Replace("{Theme_$k}", $colors[$k]) }
            $loadWin = [System.Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new($lx)))
            $loadWin.Owner = $Window
            Show-CenteredOnOwner -ChildWindow $loadWin -OwnerWindow $Window
            $loadWin.Show()

            $email = $userObj.EmailAddress.Trim()
            $job = Start-Job -ScriptBlock {
                param($tok, $apiUrl, $emailAddr)
                try {
                    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$tok`:X"))
                    $headers = @{ "Authorization" = "Basic $encoded" }
                    # Look up requester by email
                    $encodedEmail = [Uri]::EscapeDataString($emailAddr)
                    
                    $reqId = $null
                    $isAgent = $false

                    # 1. Query the STANDARD Requesters endpoint
                    $reqRes = Invoke-RestMethod -Uri "$apiUrl/api/v2/requesters?email=$encodedEmail" -Headers $headers -Method Get -ErrorAction SilentlyContinue

                    if ($reqRes -and $reqRes.requesters -and $reqRes.requesters.Count -gt 0) {
                        $reqId = $reqRes.requesters[0].id
                    } 
                    else {
                        # 2. Fallback to the AGENTS endpoint
                        $agentRes = Invoke-RestMethod -Uri "$apiUrl/api/v2/agents?email=$encodedEmail" -Headers $headers -Method Get -ErrorAction SilentlyContinue
                        if ($agentRes -and $agentRes.agents -and $agentRes.agents.Count -gt 0) {
                            $reqId = $agentRes.agents[0].id
                            $isAgent = $true
                        }
                    }

                    if ($null -eq $reqId) {
                        return [PSCustomObject]@{ Success=$true; Tickets=@(); Note="No Requester or Agent record found in Freshservice for $emailAddr." }
                    }

                    # Fetch agents and groups once for name resolution
                    $agentMap = @{}
                    try {
                        $agRes = Invoke-RestMethod -Uri "$apiUrl/api/v2/agents?per_page=100" -Headers $headers -Method Get -ErrorAction SilentlyContinue
                        if ($agRes.agents) { foreach ($a in $agRes.agents) { $agentMap[[string]$a.id] = "$($a.first_name) $($a.last_name)".Trim() } }
                    } catch {}

                    $groupMap = @{}
                    try {
                        $grRes = Invoke-RestMethod -Uri "$apiUrl/api/v2/groups?per_page=100" -Headers $headers -Method Get -ErrorAction SilentlyContinue
                        if ($grRes.groups) { foreach ($g in $grRes.groups) { $groupMap[[string]$g.id] = $g.name } }
                    } catch {}

                    # Fetch last 8 tickets for this requester (Agents and Requesters both use requester_id)
                    $tktRes = Invoke-RestMethod -Uri "$apiUrl/api/v2/tickets?requester_id=$reqId&per_page=8" -Headers $headers -Method Get -ErrorAction Stop
                    $tickets = @($tktRes.tickets | ForEach-Object {
                        $statusMap = @{2="Open";3="Pending";4="Resolved";5="Closed"}
                        $prioMap   = @{1="Low";2="Medium";3="High";4="Urgent"}
                        $assignee  = if ($_.responder_id -and $agentMap[[string]$_.responder_id]) { $agentMap[[string]$_.responder_id] } else { "Unassigned" }
                        $group     = if ($_.group_id     -and $groupMap[[string]$_.group_id])     { $groupMap[[string]$_.group_id]     } else { "" }
                        [PSCustomObject]@{
                            Id       = $_.id
                            Subject  = $_.subject
                            Status   = if ($statusMap[$_.status])   { $statusMap[$_.status]   } else { "Unknown" }
                            Priority = if ($prioMap[$_.priority])   { $prioMap[$_.priority]   } else { "?" }
                            Assignee = $assignee
                            Group    = $group
                            Created  = try { [datetime]$_.created_at } catch { $null }
                        }
                    })
                    
                    $noteStr = ""
                    if ($tickets.Count -eq 0) {
                        $agentStr = if ($isAgent) { " (User is an Agent)" } else { "" }
                        $noteStr = "No recent tickets found for $emailAddr$agentStr."
                    }
                    
                    return [PSCustomObject]@{ Success=$true; Tickets=$tickets; Note=$noteStr }
                } catch {
                    return [PSCustomObject]@{ Success=$false; ErrorMessage=$_.Exception.Message }
                }
            } -ArgumentList $token, $url, $email

            $startTime = Get-Date
            $feedTimer = New-Object System.Windows.Threading.DispatcherTimer
            $feedTimer.Interval = [TimeSpan]::FromMilliseconds(400)

            $feedTick = {
                if ($job.State -ne 'Running' -or ((Get-Date)-$startTime).TotalSeconds -ge 20) {
                    $feedTimer.Stop()
                    $loadWin.Close()
                    $result = Receive-Job $job -ErrorAction SilentlyContinue
                    Remove-Job $job -Force -ErrorAction SilentlyContinue

                    if (-not $result -or -not $result.Success) {
                        $msg = if ($result) { $result.ErrorMessage } else { "No response." }
                        Show-AppMessageBox -Message "Failed to load tickets:`n$msg" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                        return
                    }

                    # Build ticket feed window
                    $feedXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Freshservice Tickets" Width="640" Height="600" MinWidth="480" MinHeight="340"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="CanResize"
        FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Template"><Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="6"/>
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
            <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="20,18,20,12" Cursor="Hand" BorderThickness="0,0,0,1" BorderBrush="{Theme_BtnBorder}">
                <Grid>
                    <StackPanel>
                        <TextBlock Text="Recent Freshservice Tickets" FontSize="16" FontWeight="SemiBold" Foreground="{Theme_Fg}"/>
                        <TextBlock x:Name="lblFeedSub" Text="" FontSize="12" Foreground="{Theme_SecFg}" Margin="0,3,0,0"/>
                    </StackPanel>
                    <TextBlock Text="&#x2715;" HorizontalAlignment="Right" VerticalAlignment="Center"
                               Foreground="{Theme_SecFg}" FontSize="16" Cursor="Hand"
                               x:Name="btnXClose" ToolTip="Close"/>
                </Grid>
            </Border>
            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,4,0">
                <StackPanel x:Name="pnlTickets" Margin="20,14,16,8"/>
            </ScrollViewer>
            <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="12,10" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}">
                <Grid>
                    <Button x:Name="btnCloseFeed" Content="Close" Width="80" Height="28" HorizontalAlignment="Right" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" IsCancel="True"/>
                    <Thumb x:Name="resizeGrip" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="12" Height="12" Cursor="SizeNWSE" Background="Transparent" Margin="0,0,-2,-2"/>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
"@
                    $fx = $feedXaml; foreach ($k in $colors.Keys) { $fx = $fx.Replace("{Theme_$k}", $colors[$k]) }
                    $feedWin = [System.Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new($fx)))
                    $feedWin.Owner = $Window
                    # Wire X close FIRST using Preview event so it fires before DragMove captures the mouse
                    $xClose = $feedWin.FindName("btnXClose")
                    if ($xClose) {
                        $xClose.Add_PreviewMouseLeftButtonDown({
                            param($s,$e)
                            $e.Handled = $true
                            $feedWin.Close()
                        }.GetNewClosure())
                    }
                    $feedWin.FindName("TitleBar").Add_MouseLeftButtonDown({ param($s,$e); if ($e.ButtonState -eq 'Pressed') { $feedWin.DragMove() } }.GetNewClosure())
                    $feedWin.FindName("lblFeedSub").Text = "$($userObj.DisplayName) ($email)"
                    $feedWin.FindName("resizeGrip").Add_DragDelta({
                        param($s,$e)
                        $nw = [math]::Max(480, $feedWin.Width  + $e.HorizontalChange)
                        $nh = [math]::Max(340, $feedWin.Height + $e.VerticalChange)
                        $feedWin.Width = $nw; $feedWin.Height = $nh
                    }.GetNewClosure())

                    $pnl = $feedWin.FindName("pnlTickets")

                    # Status color map
                    $statusColors = @{ "Open"="#3B82F6"; "Pending"="#F59E0B"; "Resolved"="#22C55E"; "Closed"="#94A3B8" }
                    $prioColors   = @{ "Urgent"="#EF4444"; "High"="#F59E0B"; "Medium"="#3B82F6"; "Low"="#94A3B8" }

                    if ($result.Note -and $result.Tickets.Count -eq 0) {
                        $noteTb = New-Object System.Windows.Controls.TextBlock
                        $noteTb.Text = $result.Note
                        $noteTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.SecFg))
                        $noteTb.FontSize = 13; $noteTb.Margin = [System.Windows.Thickness]::new(0,4,0,4)
                        $pnl.Children.Add($noteTb) | Out-Null
                    }

                    foreach ($tkt in $result.Tickets) {
                        # Snapshot loop variables for closure (PS5.1 foreach shares scope across iterations)
                        $tktId      = $tkt.Id
                        $tktUrl     = "$url/a/tickets/$tktId"
                        $tktSubject  = $tkt.Subject
                        $tktStatus   = $tkt.Status
                        $tktPriority = $tkt.Priority
                        $tktAssignee = $tkt.Assignee
                        $tktGroup    = $tkt.Group
                        $tktCreated  = $tkt.Created

                        # Entire card is a Button so clicking anywhere opens the ticket.
                        # Custom ControlTemplate built entirely in code -- XamlReader.Load cannot
                        # parse a bare ControlTemplate fragment (it needs a full document root).
                        $card = New-Object System.Windows.Controls.Button
                        $card.Margin              = [System.Windows.Thickness]::new(0,0,0,8)
                        $card.Padding             = [System.Windows.Thickness]::new(14,10,14,10)
                        $card.Background          = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.BtnBg))
                        $card.BorderBrush         = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.GridBorder))
                        $card.BorderThickness     = [System.Windows.Thickness]::new(1)
                        $card.HorizontalAlignment = "Stretch"
                        $card.HorizontalContentAlignment = "Stretch"
                        $card.Cursor              = [System.Windows.Input.Cursors]::Hand
                        $card.ToolTip             = "Open ticket #$tktId in Freshservice"
                        $capturedUrl = $tktUrl
                        $card.Add_Click({ Start-Process "msedge.exe" -ArgumentList @("--app=$capturedUrl") }.GetNewClosure())

                        # Build ControlTemplate in code: Border wrapping a ContentPresenter
                        $bdrFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Border])
                        $bdrFactory.Name = "bd"
                        $bdrFactory.SetValue([System.Windows.Controls.Border]::CornerRadiusProperty,
                            [System.Windows.CornerRadius]::new(6))
                        # TemplateBinding equivalents via SetBinding with TemplatedParent RelativeSource
                        $tpSrc = New-Object System.Windows.Data.RelativeSource([System.Windows.Data.RelativeSourceMode]::TemplatedParent)
                        foreach ($prop in @(
                            @{DP=[System.Windows.Controls.Border]::BackgroundProperty;      Path="Background"},
                            @{DP=[System.Windows.Controls.Border]::BorderBrushProperty;     Path="BorderBrush"},
                            @{DP=[System.Windows.Controls.Border]::BorderThicknessProperty; Path="BorderThickness"},
                            @{DP=[System.Windows.Controls.Border]::PaddingProperty;         Path="Padding"}
                        )) {
                            $b = New-Object System.Windows.Data.Binding($prop.Path)
                            $b.RelativeSource = $tpSrc
                            $bdrFactory.SetBinding($prop.DP, $b)
                        }

                        $cpFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.ContentPresenter])
                        $cpFactory.SetValue([System.Windows.FrameworkElement]::HorizontalAlignmentProperty,
                            [System.Windows.HorizontalAlignment]::Stretch)
                        $cpFactory.SetValue([System.Windows.FrameworkElement]::VerticalAlignmentProperty,
                            [System.Windows.VerticalAlignment]::Top)
                        $bdrFactory.AppendChild($cpFactory)

                        $tmpl = New-Object System.Windows.Controls.ControlTemplate([System.Windows.Controls.Button])
                        $tmpl.VisualTree = $bdrFactory

                        # Hover trigger -- lighten border and shift background
                        $hoverTrigger = New-Object System.Windows.Trigger
                        $hoverTrigger.Property = [System.Windows.UIElement]::IsMouseOverProperty
                        $hoverTrigger.Value    = $true
                        $hoverSetter1 = New-Object System.Windows.Setter
                        $hoverSetter1.TargetName = "bd"
                        $hoverSetter1.Property   = [System.Windows.Controls.Border]::BackgroundProperty
                        $hoverSetter1.Value       = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.HoverBg))
                        $hoverSetter2 = New-Object System.Windows.Setter
                        $hoverSetter2.TargetName = "bd"
                        $hoverSetter2.Property   = [System.Windows.Controls.Border]::BorderBrushProperty
                        $hoverSetter2.Value       = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.PrimaryBg))
                        $hoverTrigger.Setters.Add($hoverSetter1) | Out-Null
                        $hoverTrigger.Setters.Add($hoverSetter2) | Out-Null
                        $tmpl.Triggers.Add($hoverTrigger) | Out-Null

                        # Pressed trigger -- dim slightly
                        $pressTrigger = New-Object System.Windows.Trigger
                        $pressTrigger.Property = [System.Windows.Controls.Primitives.ButtonBase]::IsPressedProperty
                        $pressTrigger.Value    = $true
                        $pressSetter = New-Object System.Windows.Setter
                        $pressSetter.TargetName = "bd"
                        $pressSetter.Property   = [System.Windows.UIElement]::OpacityProperty
                        $pressSetter.Value       = [double]0.72
                        $pressTrigger.Setters.Add($pressSetter) | Out-Null
                        $tmpl.Triggers.Add($pressTrigger) | Out-Null

                        $card.Template = $tmpl

                        $cardGrid = New-Object System.Windows.Controls.Grid
                        $col0 = New-Object System.Windows.Controls.ColumnDefinition; $col0.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
                        $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = [System.Windows.GridLength]::new(80)
                        $cardGrid.ColumnDefinitions.Add($col0); $cardGrid.ColumnDefinitions.Add($col1)

                        # Left column: ticket number (plain label), subject, date, assignee, group
                        $subjectStack = New-Object System.Windows.Controls.StackPanel

                        # Ticket number as plain accent-colored label (no longer a separate button)
                        $idTb = New-Object System.Windows.Controls.TextBlock
                        $idTb.Text       = "#$tktId"
                        $idTb.FontSize   = 11
                        $idTb.FontWeight = "SemiBold"
                        $idTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.PrimaryBg))
                        $idTb.Margin     = [System.Windows.Thickness]::new(0,0,0,2)
                        $subjectStack.Children.Add($idTb) | Out-Null

                        $subjTb = New-Object System.Windows.Controls.TextBlock
                        $subjTb.Text = $tktSubject
                        $subjTb.FontSize = 13; $subjTb.TextWrapping = "Wrap"
                        $subjTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.Fg))
                        $subjTb.Margin = [System.Windows.Thickness]::new(0,0,0,4)
                        $subjectStack.Children.Add($subjTb) | Out-Null

                        if ($tktCreated) {
                            $dateTb = New-Object System.Windows.Controls.TextBlock
                            $dateTb.Text = $tktCreated.ToString("MM/dd/yyyy")
                            $dateTb.FontSize = 11
                            $dateTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.SecFg))
                            $subjectStack.Children.Add($dateTb) | Out-Null
                        }

                        # Assignee
                        $metaLine = New-Object System.Windows.Controls.StackPanel
                        $metaLine.Orientation = "Horizontal"
                        $metaLine.Margin = [System.Windows.Thickness]::new(0,5,0,0)

                        $assigneeLbl = New-Object System.Windows.Controls.TextBlock
                        $assigneeLbl.Text = "Assignee: "
                        $assigneeLbl.FontSize = 11; $assigneeLbl.FontWeight = "SemiBold"
                        $assigneeLbl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.SecFg))
                        $metaLine.Children.Add($assigneeLbl) | Out-Null

                        $assigneeVal = New-Object System.Windows.Controls.TextBlock
                        $assigneeVal.Text = $tktAssignee
                        $assigneeVal.FontSize = 11
                        $assigneeVal.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.Fg))
                        $metaLine.Children.Add($assigneeVal) | Out-Null

                        $subjectStack.Children.Add($metaLine) | Out-Null

                        # Group (only show if populated)
                        if (-not [string]::IsNullOrWhiteSpace($tktGroup)) {
                            $groupLine = New-Object System.Windows.Controls.StackPanel
                            $groupLine.Orientation = "Horizontal"
                            $groupLine.Margin = [System.Windows.Thickness]::new(0,2,0,0)

                            $groupLbl = New-Object System.Windows.Controls.TextBlock
                            $groupLbl.Text = "Group: "
                            $groupLbl.FontSize = 11; $groupLbl.FontWeight = "SemiBold"
                            $groupLbl.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.SecFg))
                            $groupLine.Children.Add($groupLbl) | Out-Null

                            $groupVal = New-Object System.Windows.Controls.TextBlock
                            $groupVal.Text = $tktGroup
                            $groupVal.FontSize = 11
                            $groupVal.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($colors.Fg))
                            $groupLine.Children.Add($groupVal) | Out-Null

                            $subjectStack.Children.Add($groupLine) | Out-Null
                        }

                        # Status + Priority badges stacked
                        $badgeStack = New-Object System.Windows.Controls.StackPanel
                        $badgeStack.HorizontalAlignment = "Right"
                        $badgeStack.VerticalAlignment   = "Top"

                        foreach ($badge in @(@{Text=$tktStatus; Map=$statusColors}, @{Text=$tktPriority; Map=$prioColors})) {
                            $col = if ($badge.Map[$badge.Text]) { $badge.Map[$badge.Text] } else { $colors.SecFg }
                            $b = New-Object System.Windows.Controls.Border
                            $b.CornerRadius = [System.Windows.CornerRadius]::new(10)
                            $b.Padding = [System.Windows.Thickness]::new(8,3,8,3)
                            $b.Margin = [System.Windows.Thickness]::new(0,0,0,4)
                            $b.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($col))
                            $bt = New-Object System.Windows.Controls.TextBlock
                            $bt.Text = $badge.Text; $bt.FontSize = 11; $bt.FontWeight = "SemiBold"
                            $bt.Foreground = [System.Windows.Media.Brushes]::White
                            $b.Child = $bt
                            $badgeStack.Children.Add($b) | Out-Null
                        }

                        [System.Windows.Controls.Grid]::SetColumn($subjectStack, 0)
                        [System.Windows.Controls.Grid]::SetColumn($badgeStack, 1)
                        $cardGrid.Children.Add($subjectStack) | Out-Null
                        $cardGrid.Children.Add($badgeStack) | Out-Null

                        $card.Content = $cardGrid
                        $pnl.Children.Add($card) | Out-Null
                    }

                    # Wire Close button
                    $feedWin.FindName("btnCloseFeed").Add_Click({ $feedWin.Close() }.GetNewClosure())

                    Show-CenteredOnOwner -ChildWindow $feedWin -OwnerWindow $Window
                    $feedWin.Show()
                }
            }.GetNewClosure()

            $feedTimer.Add_Tick($feedTick)
            $feedTimer.Start()
        }.GetNewClosure())
    }
}
Export-ModuleMember -Function *