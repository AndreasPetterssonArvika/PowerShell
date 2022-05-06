# Setup av OU mm i AD-trädet

New-ADOrganizationalUnit -Name VUXElever -Path 'OU=Test,DC=test,DC=local'
New-ADOrganizationalUnit -Name GamlaKonton -Path 'DC=test,DC=local'
New-ADOrganizationalUnit -Name Elever -Path 'OU=GamlaKonton,DC=test,DC=local'

<#
LDIFDE ska in här om det går
#>

#<#
#Nätverksshare för användarmappar
New-Item -Path \ -Name 'Storage\Elever' -ItemType Directory
New-SmbShare -FullAccess 'BUILTIN\Everyone' -Path 'C:\Storage\Elever' -Name 'Elever$'
#>