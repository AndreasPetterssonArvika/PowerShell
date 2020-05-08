$SecurePassword = ConvertTo-SecureString -String '<ClearPassword>' -AsPlainText $true

# In default install of subdomain I needed credentials for an admin account in the parent domain to add a user to the parent domain from a server in the child domain
# Adding a user to the child domain from the parent domain did not need credentials
$Credential = Get-Credential

New-ADUser -Name '<cn>' -GivenName '<givenName>' -Surname '<SN>' -DisplayName '<displayName>' -SamAccountName '<sAMAccountName>' `
-UserPrincipalName '<user@sub.domain.com>' -EmailAddress '<user@sub.domain.com>' `
-AccountPassword $SecurePassword -Enabled $true -Server '<server.sub.domain.com>' `
-Path 'OU=<OU>,OU=<DomainBaseOU>,DC=sub,DC=domain,DC=com' -Credential $Credential