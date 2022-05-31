# Gör mitt konto till Schema Admin
$adminUser='admanpe'
Add-ADPrincipalGroupMembership -Identity $adminUser -MemberOf 'Schema Admins'

# Registera snapin för Schema management
Start-Process regsvr32 -ArgumentList "schmmgmt.dll"

# Setup av OU mm i AD-trädet
New-ADOrganizationalUnit -Name VUXElever -Path 'OU=Test,DC=test,DC=local'
New-ADOrganizationalUnit -Name GamlaKonton -Path 'DC=test,DC=local'
New-ADOrganizationalUnit -Name Elever -Path 'OU=GamlaKonton,DC=test,DC=local'

#<#
# Mappar för PowerShell-moduler
New-Item -ItemType Directory -Path C:\users\admanpe\Documents -Name 'WindowsPowerShell\Modules\Arvika-ANCUsers'
New-Item -ItemType Directory -Path C:\users\admanpe\Documents -Name 'PowerShell\Modules\Arvika-ANCUsers'
#>

#<#
#Nätverksshare för användarmappar
New-Item -Path \ -Name 'Storage\Elever' -ItemType Directory
New-SmbShare -FullAccess 'Everyone' -Path 'C:\Storage\Elever' -Name 'Elever$'
#>