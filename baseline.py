ignored_users = ["support", "admbrian", "guest"]

# Method 1: Get-ADUser -Filter
ps_filter = "LockedOut -eq $true"
for u in ignored_users:
    ps_filter += f" -and SamAccountName -ne '{u}'"
print("PS Filter:", ps_filter)

# Method 2: Get-ADUser -LDAPFilter
ldap_filter = "(&(objectClass=user)(objectCategory=person)(lockoutTime>=1)"
for u in ignored_users:
    ldap_filter += f"(!(sAMAccountName={u}))"
ldap_filter += ")"
print("LDAP Filter:", ldap_filter)
