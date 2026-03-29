# Mock benchmark showing how to test server-side vs client-side AD filtering
$IgnoredUsers = @("support", "admbrian", "guest")

# Method 1 (Current): Client-side filtering
# Measure-Command { Search-ADAccount -LockedOut | Where-Object { $IgnoredUsers -notcontains $_.SamAccountName } }

# Method 2 (New): Server-side filtering
$Filter = "LockedOut -eq `$true"
foreach ($User in $IgnoredUsers) {
    $Filter += " -and SamAccountName -ne '$User'"
}
# Measure-Command { Get-ADUser -Filter $Filter -Properties LockoutTime }

Write-Host "Because AD module is not available in the sandbox, we cannot run true benchmarks."
Write-Host "However, transitioning from client-side filtering (Search-ADAccount + Where-Object)"
Write-Host "to server-side filtering (Get-ADUser -Filter) significantly reduces AD query payload."
