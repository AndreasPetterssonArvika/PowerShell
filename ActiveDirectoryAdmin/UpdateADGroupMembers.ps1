# Uppdatera en grupp baserat på en lista med epost-adresser

if ($psISE) {
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    $basePath = $PSScriptRoot
}

# Läs in lista med identiteter till en dictionary
$memberDictionary = @{}
$infile = $basepath + "\members.txt"
$members = Get-Content -Path $infile

foreach ( $member in $members ) {
    $memberDictionary.Add($member,'member')
}

# Hämta nuvarande medlemmar i gruppen
$groupIdentity = '<group_identity>'
$curMembers = Get-ADGroupMember -Identity $groupIdentity | Get-ADUser -Properties mail | Select-Object -ExpandProperty mail

# Kontrollera nuvarande medlemmar mot dictionaryn med identiteter
# Ta bort de som inte finns i listan
foreach ( $curMember in $curMembers ) {
    if ( $memberDictionary.ContainsKey($curMember)) {
        # Redan medlem, gör inget
    } else {
        # Ta bort användaren ur gruppen
        Get-ADUser -LDAPFilter "(mail=$curMember)" | Remove-ADPrincipalGroupMembership -MemberOf $groupIdentity -Confirm:$false
    }
}

# Lägg till alla användare motsvarande identiteter i gruppen
foreach ( $member in $memberDictionary.Keys ) {
    Get-ADUser -LDAPFilter "(mail=$member)" | Add-ADPrincipalGroupMembership -MemberOf $groupIdentity
}