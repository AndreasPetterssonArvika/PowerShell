# Removes a sub-OU from all OU:s in a LDAP path
# The script handles OU:s protected from accidental deletion

$baseOU = '<OU>'
$targetOU = '<Name>'

$computerOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $FSKOU | Where-Object DistinguishedName -Like "OU=$targetOU,OU=*" | Select-Object -ExpandProperty DistinguishedName

foreach ($computerOU in $computerOUs) {
    Set-ADObject -Identity $computerOU -ProtectedFromAccidentalDeletion $false
    Remove-ADOrganizationalUnit -Identity $computerOU
}