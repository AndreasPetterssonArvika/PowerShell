#Creates a sub-OU in each of the OU:s in the selected LDAP Path


#$session = New-PSSession -ComputerName <DomainControllerName> -Credential Get-Credential
#Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
#Import-PSSession -Session $session -Module ActiveDirectory

$searchBase = '<OU>'

$ous =  Get-ADOrganizationalUnit -Filter * -SearchBase $searchBase -SearchScope OneLevel | Select-Object -ExpandProperty DistinguishedName
$newOUName = '<Name for new sub-OU>'
$newDescription = 'Description for new OU:s'

foreach ($ou in $ous) {
    $ldapPath = "LDAP://OU=$newOUName,$ou"
    $ldapPath
    if ( [adsi]::Exists($ldapPath)) {
        'OU exists'

    } else {
        'OU doesn`'t exist'
        New-ADOrganizationalUnit -Path $ou -Name $newOUName -Verbose -ProtectedFromAccidentalDeletion $true -Description $newDescription
    }

}