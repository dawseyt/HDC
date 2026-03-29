# ============================================================================
# ActiveDirectory.psm1 - Active Directory Querying and Management
# ============================================================================

function Get-LockedADUsers {
    param([Parameter(Mandatory)]$Config)
    try {
        if (Get-Module -Name ActiveDirectory) {
            $lockedAccounts = Search-ADAccount -LockedOut -Server $Config.GeneralSettings.DomainName -ErrorAction Stop
            
            $results = @()
            if ($lockedAccounts) {
                $rawUsers = $lockedAccounts | Where-Object {
                    $_.SamAccountName -and ($_.SamAccountName.ToLower() -notin $Config.GeneralSettings.FilteredUsers)
                } | ForEach-Object {
                    Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, LastLogonDate, LockoutTime, EmailAddress -Server $Config.GeneralSettings.DomainName
                }
                
                foreach ($u in $rawUsers) {
                     $results += [PSCustomObject]@{
                        Name = $u.SamAccountName
                        DisplayValue = $u.SamAccountName
                        Type = "User"
                        Description = $u.DisplayName
                        LockedOut = $true 
                        LastLogonDate = $u.LastLogonDate
                        IsOnline = $null
                        EmailAddress = $u.EmailAddress
                        LockoutTime = $u.LockoutTime
                        SamAccountName = $u.SamAccountName 
                        DisplayName = $u.DisplayName
                    }
                }
            }
            return $results
        }
        return @()
    } catch { 
        return @() 
    }
}

function Unlock-ADUsers {
    param($Usernames, $Config, $State, $IsAutoUnlock=$false)
    $errors = @()
    $eventType = if ($IsAutoUnlock) { "Auto Unlock Account" } else { "Unlock Account" }
    $detailsText = if ($IsAutoUnlock) { "Account automatically unlocked." } else { "Account manually unlocked." }

    foreach ($username in $Usernames) {
        try {
            Unlock-ADAccount -Identity $username -Server $Config.GeneralSettings.DomainName -ErrorAction Stop
            Add-AppLog -Event $eventType -Username $username -Details $detailsText -Status "Success" -Config $Config -State $State -Color "Green"
        } catch { 
            $errors += "Failed to unlock $username`: $($_.Exception.Message)"
            Add-AppLog -Event $eventType -Username $username -Details $_.Exception.Message -Status "Error" -Config $Config -State $State -Color "Red"
        }
    }
    return $errors
}

# ---------------------------------------------------------------------------
# Sanitize a string for safe embedding in an AD LDAP filter.
# Strips characters that have special meaning in LDAP filter expressions:
#   *  (  )  \  NUL
# This prevents filter injection when user-supplied search terms are used.
# ---------------------------------------------------------------------------
function Sanitize-ADFilterTerm {
    param([string]$Term)
    return $Term -replace '[*\(\)\\]', '' -replace '\x00', ''
}

function Search-ADUsers {
    param([string]$SearchTerm, $Config)
    try {
        if (Get-Module -Name ActiveDirectory) {
            $safe = Sanitize-ADFilterTerm -Term $SearchTerm
            if ([string]::IsNullOrWhiteSpace($safe)) { return @() }

            $users = Get-ADUser -Filter "SamAccountName -like '*$safe*' -or DisplayName -like '*$safe*'" -Properties DisplayName, LastLogonDate, Enabled, EmailAddress, LockedOut -Server $Config.GeneralSettings.DomainName
            
            $results = @()
            foreach ($u in $users) {
                $results += [PSCustomObject]@{
                    Name = $u.SamAccountName
                    DisplayValue = $u.SamAccountName
                    Type = "User"
                    Description = $u.DisplayName
                    LockedOut = $u.LockedOut
                    LastLogonDate = $u.LastLogonDate
                    IsOnline = $null
                    EmailAddress = $u.EmailAddress
                    SamAccountName = $u.SamAccountName
                    DisplayName = $u.DisplayName
                }
            }
            return $results
        }
        return @()
    } catch { return @() }
}

function Search-ADComputers {
    param([string]$SearchTerm, $Config)
    try {
        if (Get-Module -Name ActiveDirectory) {
            $safe = Sanitize-ADFilterTerm -Term $SearchTerm
            if ([string]::IsNullOrWhiteSpace($safe)) { return @() }

            $comps = Get-ADComputer -Filter "Name -like '*$safe*'" -Properties LastLogonDate, Enabled, OperatingSystem -Server $Config.GeneralSettings.DomainName
            
            $results = @()
            $ping = New-Object System.Net.NetworkInformation.Ping
            
            foreach ($c in $comps) {
                $isOnline = $false
                try {
                    # IMPROVEMENT: Reduced ping timeout from 200ms to 100ms to greatly speed up bulk searches
                    $reply = $ping.Send($c.Name, 100)
                    if ($reply.Status -eq "Success") {
                        $isOnline = $true
                    }
                } catch {}
                
                # REPLACED EMOJIS WITH SAFE ASCII TEXT
                $marker = if ($isOnline) { "[Online]" } else { "[Offline]" }
                
                $results += [PSCustomObject]@{
                    Name = $c.Name
                    DisplayValue = "$marker $($c.Name)"
                    Type = "Computer"
                    Description = $c.OperatingSystem
                    LockedOut = $null
                    LastLogonDate = $c.LastLogonDate
                    IsOnline = $isOnline
                    Enabled = $c.Enabled
                    OperatingSystem = $c.OperatingSystem
                }
            }
            return $results
        }
        return @()
    } catch { return @() }
}

function New-ComplexPassword {
    $len = 16
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
    $lower = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
    $digits = '0123456789'.ToCharArray()
    # REPLACED DOUBLE QUOTES WITH SINGLE QUOTES FOR SAFETY
    $special = '!@#$%^&*()_+-=[]{}|;:,.<>?'.ToCharArray()
    $passwordChars = @()
    $passwordChars += $upper | Get-Random
    $passwordChars += $lower | Get-Random
    $passwordChars += $digits | Get-Random
    $passwordChars += $special | Get-Random
    $allChars = $upper + $lower + $digits + $special
    1..($len - 4) | ForEach-Object { $passwordChars += $allChars | Get-Random }
    return -join ($passwordChars | Sort-Object { Get-Random })
}

function Reset-ADUserPassword {
    param($Username, $NewPassword, $Config, $State, $ChangeAtLogon = $true)
    try {
        $securePassword = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force
        Set-ADAccountPassword -Identity $Username -NewPassword $securePassword -Reset -Server $Config.GeneralSettings.DomainName
        Set-ADUser -Identity $Username -ChangePasswordAtLogon $ChangeAtLogon -Server $Config.GeneralSettings.DomainName
        Add-AppLog -Event "Password Reset" -Username $Username -Details "Password reset successfully (ForceChange=$ChangeAtLogon)." -Status "Success" -Config $Config -State $State -Color "Green"
        return $true
    } catch {
        Add-AppLog -Event "Password Reset" -Username $Username -Details "Reset failed: $($_.Exception.Message)" -Status "Error" -Config $Config -State $State -Color "Red"
        return $false
    }
}

function Get-UserDetails {
    param($Identity, $Type, $Config)
    try {
        if ($Type -eq "Computer") {
            return Get-ADComputer -Identity $Identity -Properties * -Server $Config.GeneralSettings.DomainName
        } else {
            return Get-ADUser -Identity $Identity -Properties * -Server $Config.GeneralSettings.DomainName
        }
    } catch { return $null }
}

# Explicitly export all functions so the main script can see them
Export-ModuleMember -Function *