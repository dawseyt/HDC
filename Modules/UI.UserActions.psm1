# ============================================================================
# UI.UserActions.psm1 - Active Directory User Interface Logic
# ============================================================================

function Register-UserUIEvents {
    param($Window, $Config, $State)

    $AppRoot = Split-Path -Path $PSScriptRoot -Parent
    
    $lvData = $Window.FindName("lvData")
    $btnUnlock = $Window.FindName("btnUnlock")
    $btnUnlockAll = $Window.FindName("btnUnlockAll")
    $chkAutoUnlock = $Window.FindName("chkAutoUnlock")
    
    $ctxUnlock = $Window.FindName("ctxUnlock")
    $ctxReset = $Window.FindName("ctxReset")
    $ctxDetails = $Window.FindName("ctxDetails")
    $ctxLockoutSource = $Window.FindName("ctxLockoutSource")
    $ctxGroups = $Window.FindName("ctxGroups")
    $ctxTicketFeed = $Window.FindName("ctxTicketFeed")
    $ctxLastLogonLocation = $Window.FindName("MenuItem_LastLogonLocation")
    
    $overlayReset = $Window.FindName("overlayReset")
    $txtPwdTitle = $Window.FindName("txtPwdTitle")
    $txtResetUsername = $Window.FindName("txtResetUsername")
    $txtResetPassword = $Window.FindName("txtResetPassword")
    $txtResetPasswordVisible = $Window.FindName("txtResetPasswordVisible")
    $txtConfirmPassword = $Window.FindName("txtConfirmPassword")
    $txtConfirmPasswordVisible = $Window.FindName("txtConfirmPasswordVisible")
    $chkShowPassword = $Window.FindName("chkShowPassword")
    $lblMatchStatus = $Window.FindName("lblMatchStatus")
    $chkMustChange = $Window.FindName("chkMustChange")
    $btnConfirmReset = $Window.FindName("btnConfirmReset")
    $btnCancelReset = $Window.FindName("btnCancelReset")
    $btnGeneratePass = $Window.FindName("btnGeneratePass")

    # --- Dynamic Menu Injection: Last Logon Location ---
    if ($lvData -and $lvData.ContextMenu) {
        foreach ($item in $lvData.ContextMenu.Items) {
            if ($item -is [System.Windows.Controls.MenuItem] -and ($item.Name -eq "MenuItem_LastLogonLocation" -or $item.Header -match "Find Last Logon Location")) {
                $ctxLastLogonLocation = $item; break
            }
        }

        if (-not $ctxLastLogonLocation) {
            $ctxLastLogonLocation = New-Object System.Windows.Controls.MenuItem
            $ctxLastLogonLocation.Name = "MenuItem_LastLogonLocation"
            $ctxLastLogonLocation.Header = "Find Last Logon Location"
            $lIcon = New-Object System.Windows.Controls.TextBlock
            $lIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $lIcon.Text = [char]0xE716 # Computer/Network Icon
            $lIcon.FontSize = 14
            $lIcon.FontWeight = [System.Windows.FontWeights]::Bold
            $lIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "AccentFill")
            $ctxLastLogonLocation.Icon = $lIcon
            
            $inserted = $false
            if ($ctxLockoutSource) {
                $idx = $lvData.ContextMenu.Items.IndexOf($ctxLockoutSource)
                if ($idx -ge 0) {
                    $lvData.ContextMenu.Items.Insert($idx + 1, $ctxLastLogonLocation)
                    $inserted = $true
                }
            }
            if (-not $inserted) { $lvData.ContextMenu.Items.Add($ctxLastLogonLocation) | Out-Null }
        }

        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            $isUser = ($sel -and $sel.Type -eq "User")
            $vis = if ($isUser) { "Visible" } else { "Collapsed" }
            if ($ctxUnlock)            { $ctxUnlock.Visibility            = $vis }
            if ($ctxReset)             { $ctxReset.Visibility             = $vis }
            if ($ctxDetails)           { $ctxDetails.Visibility           = $vis }
            if ($ctxLockoutSource)     { $ctxLockoutSource.Visibility     = $vis }
            if ($ctxLastLogonLocation) { $ctxLastLogonLocation.Visibility = $vis }
            if ($ctxGroups)            { $ctxGroups.Visibility            = $vis }
            if ($ctxTicketFeed)        { $ctxTicketFeed.Visibility        = $vis }
        }.GetNewClosure())
    }

    $UnlockAction = {
        if ($lvData.SelectedItems.Count -eq 0) {
            Show-AppMessageBox -Message "Please select at least one item." -Title "No Selection" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
            return
        }
        $users = @()
        foreach ($item in $lvData.SelectedItems) { if ($item.Type -eq 'User') { $users += $item.Name } }
        if ($users.Count -gt 0) {
            $confirm = Show-AppMessageBox -Message "Unlock $($users.Count) selected users?" -Title "Confirm" -ButtonType "YesNo" -IconType "Question" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
            if ($confirm -eq "Yes") {
                Unlock-ADUsers -Usernames $users -Config $Config -State $State
                if ($State.Actions.RefreshData) { & $State.Actions.RefreshData }
            }
        } else {
            Show-AppMessageBox -Message "Selected item(s) are not Users or cannot be unlocked." -Title "Info" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
        }
    }.GetNewClosure()

    if ($btnUnlock) { $btnUnlock.Add_Click($UnlockAction) }
    if ($ctxUnlock) { $ctxUnlock.Add_Click($UnlockAction) }

    if ($btnUnlockAll) {
        $btnUnlockAll.Add_Click({
            if ($lvData.Items.Count -eq 0) { return }
            $confirm = Show-AppMessageBox -Message "Are you sure you want to unlock ALL listed accounts?" -Title "Confirm Unlock All" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
            if ($confirm -eq "Yes") {
                $users = @()
                foreach ($item in $lvData.Items) { if ($item.Type -eq 'User') { $users += $item.Name } }
                Unlock-ADUsers -Usernames $users -Config $Config -State $State
                if ($State.Actions.RefreshData) { & $State.Actions.RefreshData }
            }
        }.GetNewClosure())
    }

    $State.AutoUnlockTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $State.AutoUnlockTimer.Interval = [TimeSpan]::FromSeconds($Config.AutoSettings.AutoUnlockIntervalSeconds)
    $State.AutoUnlockTimer.Add_Tick({
        if ($chkAutoUnlock -and $chkAutoUnlock.IsChecked -and $lvData -and $lvData.Items.Count -gt 0) {
            $users = @()
            foreach ($item in $lvData.Items) { if ($item.Type -eq "User") { $users += $item.Name } }
            if ($users.Count -gt 0) { 
                Unlock-ADUsers -Usernames $users -Config $Config -State $State -IsAutoUnlock $true
                if ($State.Actions.RefreshData) { & $State.Actions.RefreshData }
            }
        }
    }.GetNewClosure())
    $State.AutoUnlockTimer.Start()

    # --- Reset Password Logic ---
    if ($ctxReset) {
        $ctxReset.Add_Click({
            if ($lvData.SelectedItem.Type -eq "User") {
                if ($txtResetUsername) { $txtResetUsername.Text = $lvData.SelectedItem.Name }
                if ($txtPwdTitle) { $txtPwdTitle.Text = "Reset password for " + $lvData.SelectedItem.Name }
                if ($txtResetPassword) { $txtResetPassword.Password = "" }
                if ($txtConfirmPassword) { $txtConfirmPassword.Password = "" }
                if ($txtResetPasswordVisible) { $txtResetPasswordVisible.Text = "" }
                if ($txtConfirmPasswordVisible) { $txtConfirmPasswordVisible.Text = "" }
                if ($chkShowPassword) { $chkShowPassword.IsChecked = $false }
                if ($txtResetPassword) { $txtResetPassword.Visibility = "Visible" }
                if ($txtConfirmPassword) { $txtConfirmPassword.Visibility = "Visible" }
                if ($txtResetPasswordVisible) { $txtResetPasswordVisible.Visibility = "Collapsed" }
                if ($txtConfirmPasswordVisible) { $txtConfirmPasswordVisible.Visibility = "Collapsed" }
                if ($lblMatchStatus) { $lblMatchStatus.Text = "" }
                if ($overlayReset) { $overlayReset.Visibility = "Visible" }
            }
        }.GetNewClosure())
    }
    
    if ($btnCancelReset) { $btnCancelReset.Add_Click({ if ($overlayReset) { $overlayReset.Visibility = "Collapsed" } }.GetNewClosure()) }
    
    $UpdateMatchStatus = {
        $pass = ""; $conf = ""
        if ($chkShowPassword -and $chkShowPassword.IsChecked) {
            if ($txtResetPasswordVisible) { $pass = $txtResetPasswordVisible.Text }
            if ($txtConfirmPasswordVisible) { $conf = $txtConfirmPasswordVisible.Text }
        } else {
            if ($txtResetPassword) { $pass = $txtResetPassword.Password }
            if ($txtConfirmPassword) { $conf = $txtConfirmPassword.Password }
        }
        if ($lblMatchStatus) {
            if ([string]::IsNullOrEmpty($conf)) { $lblMatchStatus.Text = "" } 
            elseif ($pass -eq $conf) { $lblMatchStatus.Text = "$([char]0x2713) Passwords Match"; $lblMatchStatus.Foreground = [System.Windows.Media.Brushes]::Green } 
            else { $lblMatchStatus.Text = "$([char]0x2717) Passwords do not match"; $lblMatchStatus.Foreground = [System.Windows.Media.Brushes]::Red }
        }
    }.GetNewClosure()

    if ($txtResetPassword) { $txtResetPassword.Add_PasswordChanged($UpdateMatchStatus) }
    if ($txtConfirmPassword) { $txtConfirmPassword.Add_PasswordChanged($UpdateMatchStatus) }
    if ($txtResetPasswordVisible) { $txtResetPasswordVisible.Add_TextChanged($UpdateMatchStatus) }
    if ($txtConfirmPasswordVisible) { $txtConfirmPasswordVisible.Add_TextChanged($UpdateMatchStatus) }

    if ($chkShowPassword) {
        $chkShowPassword.Add_Checked({
            if ($txtResetPasswordVisible -and $txtResetPassword) { $txtResetPasswordVisible.Text = $txtResetPassword.Password }
            if ($txtConfirmPasswordVisible -and $txtConfirmPassword) { $txtConfirmPasswordVisible.Text = $txtConfirmPassword.Password }
            if ($txtResetPassword) { $txtResetPassword.Visibility = "Collapsed" }
            if ($txtConfirmPassword) { $txtConfirmPassword.Visibility = "Collapsed" }
            if ($txtResetPasswordVisible) { $txtResetPasswordVisible.Visibility = "Visible" }
            if ($txtConfirmPasswordVisible) { $txtConfirmPasswordVisible.Visibility = "Visible" }
            & $UpdateMatchStatus
        }.GetNewClosure())
        
        $chkShowPassword.Add_Unchecked({
            if ($txtResetPassword -and $txtResetPasswordVisible) { $txtResetPassword.Password = $txtResetPasswordVisible.Text }
            if ($txtConfirmPassword -and $txtConfirmPasswordVisible) { $txtConfirmPassword.Password = $txtConfirmPasswordVisible.Text }
            if ($txtResetPassword) { $txtResetPassword.Visibility = "Visible" }
            if ($txtConfirmPassword) { $txtConfirmPassword.Visibility = "Visible" }
            if ($txtResetPasswordVisible) { $txtResetPasswordVisible.Visibility = "Collapsed" }
            if ($txtConfirmPasswordVisible) { $txtConfirmPasswordVisible.Visibility = "Collapsed" }
            & $UpdateMatchStatus
        }.GetNewClosure())
    }

    if ($btnGeneratePass) {
        $btnGeneratePass.Add_Click({ 
            $newPass = New-ComplexPassword
            if ($txtResetPassword) { $txtResetPassword.Password = $newPass }
            if ($txtConfirmPassword) { $txtConfirmPassword.Password = $newPass }
            if ($txtResetPasswordVisible) { $txtResetPasswordVisible.Text = $newPass }
            if ($txtConfirmPasswordVisible) { $txtConfirmPasswordVisible.Text = $newPass }
            & $UpdateMatchStatus 
        }.GetNewClosure())
    }
    
    if ($btnConfirmReset) {
        $btnConfirmReset.Add_Click({
            $u = if ($txtResetUsername) { $txtResetUsername.Text } else { "" }
            $mustChange = if ($chkMustChange) { $chkMustChange.IsChecked } else { $true }
            
            $p = ""; $c = ""
            if ($chkShowPassword -and $chkShowPassword.IsChecked) {
                if ($txtResetPasswordVisible) { $p = $txtResetPasswordVisible.Text }
                if ($txtConfirmPasswordVisible) { $c = $txtConfirmPasswordVisible.Text }
            } else {
                if ($txtResetPassword) { $p = $txtResetPassword.Password }
                if ($txtConfirmPassword) { $c = $txtConfirmPassword.Password }
            }
        
            if ($p -ne $c) { Show-AppMessageBox -Message "Passwords do not match." -Title "Validation Error" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); return }
            if ($p.Length -lt 16 -or $p -notmatch '[A-Z]' -or $p -notmatch '[0-9]' -or $p -notmatch '[^a-zA-Z0-9]') {
                Show-AppMessageBox -Message "Password does not meet complexity requirements." -Title "Error" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); return
            }
            
            if (Reset-ADUserPassword -Username $u -NewPassword $p -Config $Config -State $State -ChangeAtLogon $mustChange) {
                Show-AppMessageBox -Message "Password reset." -Title "Success" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                if ($overlayReset) { $overlayReset.Visibility = "Collapsed" }
            }
        }.GetNewClosure())
    }

    # --- Group Membership (read + manage) ---
    if ($ctxGroups) {
        $ctxGroups.Add_Click({
            if ($lvData.SelectedItem) {
                $id   = $lvData.SelectedItem.Name
                $type = $lvData.SelectedItem.Type
                
                $colors = Get-FluentThemeColors $State
                $grpWin = Load-XamlWindow -XamlPath (Join-Path $AppRoot "UI\Windows\GroupMembership.xaml") -ThemeColors $colors
                $grpWin.Owner = $Window
                $grpWin.Title = "Group Membership - $id"

                $lblTitle     = $grpWin.FindName("lblHeaderTitle")
                $lblSub       = $grpWin.FindName("lblHeaderSub")
                $lblStatus    = $grpWin.FindName("lblGrpStatus")
                $lvGrps       = $grpWin.FindName("lvGroups")
                $txtAddGrp    = $grpWin.FindName("txtAddGroup")
                $btnSearchGrp = $grpWin.FindName("btnSearchGroup")
                $btnAddGrp    = $grpWin.FindName("btnAddGroup")
                $btnRemoveGrp = $grpWin.FindName("btnRemoveGroup")
                $lbGrpSearch  = $grpWin.FindName("lbGroupSearch")
                $borderSearch = $grpWin.FindName("borderSearchResults")
                $btnCloseGrp  = $grpWin.FindName("btnClose")

                # --- Group Membership Sort Logic ---
                $sortState = @{ Column = "Name"; Descending = $false }
                
                $SortGroupList = {
                    param($sortBy)
                    if ($lvGrps -and $lvGrps.ItemsSource) {
                        $items = @($lvGrps.ItemsSource)
                        if ($items.Count -gt 0) {
                            $sorted = $items | Sort-Object -Property $sortBy -Descending:$sortState.Descending
                            $lvGrps.ItemsSource = $null
                            $lvGrps.ItemsSource = @($sorted)
                        }
                    }
                }.GetNewClosure()

                if ($lvGrps) {
                    $lvGrps.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
                        param($sender, $e)
                        $source = $e.OriginalSource
                        
                        while ($source -and -not ($source -is [System.Windows.Controls.GridViewColumnHeader])) {
                            if ($source -is [System.Windows.FrameworkElement]) { $source = $source.Parent } else { break }
                        }
                        
                        if ($source -and ($source -is [System.Windows.Controls.GridViewColumnHeader]) -and $source.Role -ne "Padding") {
                            $column = $source.Column
                            $sortBy = if ($column.DisplayMemberBinding) { $column.DisplayMemberBinding.Path.Path } else { "Name" }
                            
                            if ($sortBy) {
                                if ($sortState.Column -eq $sortBy) { 
                                    $sortState.Descending = -not $sortState.Descending 
                                } else { 
                                    $sortState.Descending = $false
                                    $sortState.Column = $sortBy 
                                }
                                & $SortGroupList -sortBy $sortBy
                            }
                        }
                    }.GetNewClosure())
                }
                
                if ($lblTitle) { $lblTitle.Text = "Member Of" }
                if ($lblSub)   { $lblSub.Text   = $id }

                # --- Async Data Loader ---
                $LoadGroupsAsync = {
                    $grpWin.Dispatcher.Invoke({
                        if ($lblStatus) { $lblStatus.Text = "Querying Active Directory..." }
                        if ($lvGrps) { $lvGrps.ItemsSource = $null }
                        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                    })
                    
                    $action = [System.Action]{
                        try {
                            # Bypass buggy Get-ADPrincipalGroupMembership by reading MemberOf and PrimaryGroupID directly
                            $groups = @()
                            $targetObj = $null
                            
                            if ($type -eq "Computer") {
                                $targetObj = Get-ADComputer -Identity $id -Properties MemberOf, PrimaryGroupID, SID -Server $Config.GeneralSettings.DomainName -ErrorAction Stop
                            } else {
                                $targetObj = Get-ADUser -Identity $id -Properties MemberOf, PrimaryGroupID, SID -Server $Config.GeneralSettings.DomainName -ErrorAction Stop
                            }

                            if ($targetObj.MemberOf) {
                                foreach ($dn in $targetObj.MemberOf) {
                                    try {
                                        $grp = Get-ADGroup -Identity $dn -Server $Config.GeneralSettings.DomainName -ErrorAction Stop | 
                                               Select-Object Name, GroupScope, GroupCategory, DistinguishedName
                                        if ($grp) { $groups += $grp }
                                    } catch {}
                                }
                            }

                            if ($targetObj.PrimaryGroupID -and $targetObj.SID) {
                                try {
                                    $domainSID = $targetObj.SID.Value -replace '-\d+$', ''
                                    $primaryGroupSID = "$domainSID-$($targetObj.PrimaryGroupID)"
                                    $primaryGrp = Get-ADGroup -Identity $primaryGroupSID -Server $Config.GeneralSettings.DomainName -ErrorAction Stop | 
                                                  Select-Object Name, GroupScope, GroupCategory, DistinguishedName
                                    if ($primaryGrp) { $groups += $primaryGrp }
                                } catch {}
                            }
                            
                            $groups = $groups | Sort-Object Name -Unique
                            
                            $grpWin.Dispatcher.Invoke({
                                if ($lvGrps) { 
                                    $lvGrps.ItemsSource = $groups 
                                    & $SortGroupList -sortBy $sortState.Column
                                }
                                if ($lblTitle) { $lblTitle.Text = "Member Of ($($groups.Count) Groups)" }
                                if ($lblStatus) { $lblStatus.Text = "Found $($groups.Count) group(s)." }
                                if ($btnRemoveGrp) { $btnRemoveGrp.IsEnabled = $false }
                                [System.Windows.Input.Mouse]::OverrideCursor = $null
                            })
                        } catch {
                            $err = $_.Exception.Message
                            $grpWin.Dispatcher.Invoke({
                                if ($lblStatus) { $lblStatus.Text = "Error: $err"; $lblStatus.Foreground = [System.Windows.Media.Brushes]::Red }
                                [System.Windows.Input.Mouse]::OverrideCursor = $null
                                Show-AppMessageBox -Message "Failed to fetch groups for $id.`n`n$err" -Title "Active Directory Error" -IconType "Error" -OwnerWindow $grpWin -ThemeColors $colors | Out-Null
                            })
                        }
                    }
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $action) | Out-Null
                }.GetNewClosure()

                & $LoadGroupsAsync

                if ($lvGrps) {
                    $lvGrps.Add_SelectionChanged({
                        if ($btnRemoveGrp) { $btnRemoveGrp.IsEnabled = ($null -ne $lvGrps.SelectedItem) }
                    }.GetNewClosure())
                }

                if ($btnRemoveGrp) {
                    $btnRemoveGrp.Add_Click({
                        $selGroup = $lvGrps.SelectedItem
                        if (-not $selGroup) { return }
                        $conf = Show-AppMessageBox -Message "Remove $id from group `'$($selGroup.Name)`'?" -Title "Confirm Remove" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $grpWin -ThemeColors $colors
                        if ($conf -ne "Yes") { return }
                        try {
                            if ($type -eq "Computer") {
                                Remove-ADGroupMember -Identity $selGroup.DistinguishedName -Members (Get-ADComputer -Identity $id -Server $Config.GeneralSettings.DomainName) -Server $Config.GeneralSettings.DomainName -Confirm:$false -ErrorAction Stop
                            } else {
                                Remove-ADGroupMember -Identity $selGroup.DistinguishedName -Members (Get-ADUser -Identity $id -Server $Config.GeneralSettings.DomainName) -Server $Config.GeneralSettings.DomainName -Confirm:$false -ErrorAction Stop
                            }
                            if ($lblStatus) { $lblStatus.Foreground = [System.Windows.Media.Brushes]::Green; $lblStatus.Text = "Removed from `'$($selGroup.Name)`'." }
                            Add-AppLog -Event "Group Remove" -Username $id -Details "Removed from group: $($selGroup.Name)" -Config $Config -State $State -Status "Success" -Color "Green"
                            & $LoadGroupsAsync
                        } catch {
                            if ($lblStatus) { $lblStatus.Foreground = [System.Windows.Media.Brushes]::Red; $lblStatus.Text = "Error: $($_.Exception.Message)" }
                            Add-AppLog -Event "Group Remove" -Username $id -Details "Failed: $($_.Exception.Message)" -Config $Config -State $State -Status "Error" -Color "Red"
                        }
                    }.GetNewClosure())
                }

                $SearchGroups = {
                    $term = if ($txtAddGrp) { $txtAddGrp.Text.Trim() } else { "" }
                    if ([string]::IsNullOrWhiteSpace($term)) { return }
                    try {
                        $safe = $term
                        foreach ($ch in @('*', '(', ')', '/', '[', ']')) { $safe = [string]::new($safe).Replace($ch, '') }
                        if ([string]::IsNullOrWhiteSpace($safe)) { if ($lblStatus) { $lblStatus.Text = "Enter a group name to search." }; return }
                        
                        $found = Get-ADGroup -Filter "Name -like `'*$safe*`'" -Server $Config.GeneralSettings.DomainName -Properties Name, GroupScope, GroupCategory | Select-Object -First 30 Name, GroupScope, GroupCategory, DistinguishedName | Sort-Object Name
                        if ($lbGrpSearch) {
                            $lbGrpSearch.Items.Clear()
                            foreach ($g in $found) { $lbGrpSearch.Items.Add($g.Name) | Out-Null }
                            $lbGrpSearch.Tag = $found
                        }
                        if ($borderSearch) { $borderSearch.Visibility = if ($found.Count -gt 0) { "Visible" } else { "Collapsed" } }
                        if ($lblStatus)    { $lblStatus.Text = "Found $($found.Count) group(s)." }
                        if ($btnAddGrp)    { $btnAddGrp.IsEnabled = $false }
                    } catch {
                        if ($lblStatus) { $lblStatus.Text = "Search failed: $($_.Exception.Message)" }
                    }
                }.GetNewClosure()

                if ($btnSearchGrp) { $btnSearchGrp.Add_Click($SearchGroups) }
                if ($txtAddGrp)    { $txtAddGrp.Add_KeyDown({ param($s,$e); if ($e.Key -eq "Enter") { & $SearchGroups } }.GetNewClosure()) }

                if ($lbGrpSearch) {
                    $lbGrpSearch.Add_SelectionChanged({
                        if ($btnAddGrp) { $btnAddGrp.IsEnabled = ($null -ne $lbGrpSearch.SelectedItem) }
                    }.GetNewClosure())
                }

                if ($btnAddGrp) {
                    $btnAddGrp.Add_Click({
                        if (-not $lbGrpSearch.SelectedItem) { return }
                        $selName  = $lbGrpSearch.SelectedItem
                        $allFound = $lbGrpSearch.Tag
                        $selGroup = @($allFound | Where-Object { $_.Name -eq $selName }) | Select-Object -First 1
                        if (-not $selGroup) { return }
                        $conf = Show-AppMessageBox -Message "Add $id to group `'$($selGroup.Name)`'?" -Title "Confirm Add" -ButtonType "YesNo" -IconType "Question" -OwnerWindow $grpWin -ThemeColors $colors
                        if ($conf -ne "Yes") { return }
                        try {
                            if ($type -eq "Computer") {
                                Add-ADGroupMember -Identity $selGroup.DistinguishedName -Members (Get-ADComputer -Identity $id -Server $Config.GeneralSettings.DomainName) -Server $Config.GeneralSettings.DomainName -Confirm:$false -ErrorAction Stop
                            } else {
                                Add-ADGroupMember -Identity $selGroup.DistinguishedName -Members (Get-ADUser -Identity $id -Server $Config.GeneralSettings.DomainName) -Server $Config.GeneralSettings.DomainName -Confirm:$false -ErrorAction Stop
                            }
                            if ($lblStatus) { $lblStatus.Foreground = [System.Windows.Media.Brushes]::Green; $lblStatus.Text = "Added to `'$($selGroup.Name)`'." }
                            Add-AppLog -Event "Group Add" -Username $id -Details "Added to group: $($selGroup.Name)" -Config $Config -State $State -Status "Success" -Color "Green"
                            if ($txtAddGrp)    { $txtAddGrp.Text = "" }
                            if ($lbGrpSearch)  { $lbGrpSearch.Items.Clear() }
                            if ($borderSearch) { $borderSearch.Visibility = "Collapsed" }
                            if ($btnAddGrp)    { $btnAddGrp.IsEnabled = $false }
                            & $LoadGroupsAsync
                        } catch {
                            if ($lblStatus) { $lblStatus.Foreground = [System.Windows.Media.Brushes]::Red; $lblStatus.Text = "Error: $($_.Exception.Message)" }
                            Add-AppLog -Event "Group Add" -Username $id -Details "Failed: $($_.Exception.Message)" -Config $Config -State $State -Status "Error" -Color "Red"
                        }
                    }.GetNewClosure())
                }

                if ($btnCloseGrp) { $btnCloseGrp.Add_Click({ $grpWin.Close() }.GetNewClosure()) }

                Show-CenteredOnOwner -ChildWindow $grpWin -OwnerWindow $Window
                $grpWin.ShowDialog() | Out-Null
            }
        }.GetNewClosure())
    }

    # --- Lockout Source Finder ---
    if ($ctxLockoutSource) {
        $ctxLockoutSource.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $user = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $loadingXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="320" Height="120" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                            <TextBlock Text="Searching PDC for Lockout Source..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                            <ProgressBar IsIndeterminate="True" Width="240" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                        </StackPanel>
                    </Border>
                </Window>
"@
                $xamlText = $loadingXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 

                Add-AppLog -Event "Query" -Username "System" -Details "Searching PDC for lockout source of $user..." -Config $Config -State $State -Status "Info"

                $job = Start-Job -ScriptBlock {
                    param($u)
                    try {
                        $pdc = (Get-ADDomain).PDCEmulator
                        if (-not $pdc) { throw "Could not determine PDC Emulator." }
                        $event = Get-WinEvent -ComputerName $pdc -FilterHashtable @{LogName='Security'; Id=4740; Data=$u} -MaxEvents 1 -ErrorAction Stop
                        if ($event) {
                            $callerPC = "Unknown"
                            if ($event.Message -match 'Caller Computer Name:\s+([^\r\n]+)') { $callerPC = $matches[1].Trim() }
                            if ($callerPC -eq "Unknown" -or [string]::IsNullOrWhiteSpace($callerPC)) {
                                $xml = [xml]$event.ToXml(); $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $nsmgr.AddNamespace("ns", "http://schemas.microsoft.com/win/2004/08/events/event")
                                $node = $xml.SelectSingleNode("//ns:Data[@Name='CallerComputerName']", $nsmgr)
                                if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) { $callerPC = $node.InnerText.Trim() }
                            }
                            return [PSCustomObject]@{ Found = $true; CallerComputer = $callerPC; Time = $event.TimeCreated; PDC = $pdc }
                        } else { return [PSCustomObject]@{ Found = $false; PDC = $pdc } }
                    } catch { throw $_.Exception.Message }
                } -ArgumentList $user

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                $timerTick = {
                    if ($job.State -ne 'Running') {
                        $timer.Stop(); $loadWin.Close()
                        if ($job.State -eq 'Completed') {
                            $result = Receive-Job $job -ErrorAction SilentlyContinue
                            if ($result.Found) { Show-AppMessageBox -Message "Lockout Source Found!`n`nUser: $user`nSource Computer: $($result.CallerComputer)`nTime: $($result.Time)`n`nDomain Controller: $($result.PDC)" -Title "Lockout Source" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors } 
                            else { Show-AppMessageBox -Message "No recent lockout events (Event ID 4740) found for $user on the PDC ($($result.PDC))." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors }
                        } else {
                            $reason = if ($job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Unknown error" }
                            Show-AppMessageBox -Message "Failed to query lockout source:`n$reason" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors
                        }
                        Remove-Job $job -Force
                    }
                }.GetNewClosure()
                $timer.Add_Tick($timerTick)
                $timer.Start()
            }
        }.GetNewClosure())
    }

    # --- Find Last Logon Location ---
    if ($ctxLastLogonLocation) {
        $ctxLastLogonLocation.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $user = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $loadingXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="380" Height="120" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                            <TextBlock Text="Searching PDC for Last Logon Location..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                            <ProgressBar IsIndeterminate="True" Width="300" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                        </StackPanel>
                    </Border>
                </Window>
"@
                $xamlText = $loadingXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 

                Add-AppLog -Event "Query" -Username "System" -Details "Searching PDC for last logon location of $user..." -Config $Config -State $State -Status "Info"

                $job = Start-Job -ScriptBlock {
                    param($u)
                    try {
                        $pdc = (Get-ADDomain).PDCEmulator
                        if (-not $pdc) { throw "Could not determine PDC Emulator." }
                        
                        # Pull top 50 Logon events (4624) for this user
                        $events = Get-WinEvent -ComputerName $pdc -FilterHashtable @{LogName='Security'; Id=4624; Data=$u} -MaxEvents 50 -ErrorAction Stop
                        
                        foreach ($evt in $events) {
                            # Extract IpAddress safely using the Event XML schema
                            $xml = [xml]$evt.ToXml()
                            $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
                            $nsmgr.AddNamespace("ns", "http://schemas.microsoft.com/win/2004/08/events/event")
                            
                            $ipNode = $xml.SelectSingleNode("//ns:Data[@Name='IpAddress']", $nsmgr)
                            $ip = if ($ipNode) { $ipNode.InnerText.Trim() } else { "" }
                            
                            # Filter out local loopback or blank IPs
                            if (-not [string]::IsNullOrWhiteSpace($ip) -and $ip -ne "-" -and $ip -ne "::1" -and $ip -ne "127.0.0.1") {
                                $hostName = "Unknown"
                                try { 
                                    # Reverse DNS lookup to map the IP to the Workstation Name
                                    $hostEntry = [System.Net.Dns]::GetHostEntry($ip)
                                    if ($hostEntry -and $hostEntry.HostName) { $hostName = $hostEntry.HostName }
                                } catch {}
                                
                                return [PSCustomObject]@{ Found = $true; IPAddress = $ip; HostName = $hostName; Time = $evt.TimeCreated; PDC = $pdc }
                            }
                        }
                        return [PSCustomObject]@{ Found = $false; PDC = $pdc }
                    } catch { throw $_.Exception.Message }
                } -ArgumentList $user

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                $timerTick = {
                    if ($job.State -ne 'Running') {
                        $timer.Stop(); $loadWin.Close()
                        if ($job.State -eq 'Completed') {
                            $result = Receive-Job $job -ErrorAction SilentlyContinue
                            if ($result.Found) { 
                                Show-AppMessageBox -Message "Last Logon Location Found!`n`nUser: $user`nHostname: $($result.HostName)`nIP Address: $($result.IPAddress)`nTime: $($result.Time)`n`nDomain Controller: $($result.PDC)" -Title "Last Logon Location" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                            } else { 
                                Show-AppMessageBox -Message "No recent remote/network logon events (Event ID 4624) found for $user on the PDC ($($result.PDC))." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                            }
                        } else {
                            $reason = if ($job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Unknown error" }
                            Show-AppMessageBox -Message "Failed to query last logon location:`n$reason" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                        }
                        Remove-Job $job -Force
                    }
                }.GetNewClosure()
                $timer.Add_Tick($timerTick)
                $timer.Start()
            }
        }.GetNewClosure())
    }
}
Export-ModuleMember -Function *