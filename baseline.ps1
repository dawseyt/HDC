$IgnoredUsers = @("support", "admbrian", "guest")

# Mock the current client-side filter string
$ldapFilter = "(&(objectClass=user)(objectCategory=person)(lockoutTime>=1)"
foreach ($user in $IgnoredUsers) {
    $ldapFilter += "(!(sAMAccountName=$user))"
}
$ldapFilter += ")"

Write-Host "Constructed LDAP Filter: $ldapFilter"

# Alternative PowerShell Filter
$psFilter = "LockedOut -eq `$true"
foreach ($user in $IgnoredUsers) {
    $psFilter += " -and SamAccountName -ne '$user'"
}

Write-Host "Constructed PS Filter: $psFilter"
