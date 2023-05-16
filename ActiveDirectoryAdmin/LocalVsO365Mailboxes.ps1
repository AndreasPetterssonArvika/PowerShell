# Kontrollera mailboxar
# Kategorier av intresse
# - de som har ett värde i msExchMailboxGuid, tolkat som att de har en mailbox
# - de som har ett värde i msExchMailboxGuid men msexchRemoteRecipientType skild från 4, tolkat som att de har en lokal mailbox
# - de som har ett värde i msExchMailboxGuid och msexchRemoteRecipientType som är 4, tolkat som att de har en Office-365-mailbox
#
# Skriptet räknar också de användare som är låsta, eftersom de i de flesta fall är funktionskonton som inte loggar in utan
# används av andra användare.

[cmdletbinding()]
param(
    [switch]$ListMailboxes
)

$now = Get-Date -Format "yyMMdd_HHmm"

$outfile = "LocalVsO365Mailboxes_$now.txt"

$ldapHasMailbox = '(msExchMailboxGuid=*)'
$ldapHasLocalMailbox = '(&(msExchMailboxGuid=*)(!(msexchRemoteRecipientType=4)))'
$ldapHasO365Mailbox = '(&(msExchMailboxGuid=*)(msexchRemoteRecipientType=4))'

# Attribut som behöver definieras utöver de som Get-ADUser levererar
$attributeList = @('mail','msExchMailboxGuid','msexchRemoteRecipientType','Enabled')

# Räkna användare med mailbox
$usersWithMailbox = Get-ADUser -LDAPFilter $ldapHasMailbox -Properties $attributeList
$numTotalMailboxes = $usersWithMailbox | Measure-Object | Select-Object -ExpandProperty Count

# Räkna användare med lokal mailbox
$usersWithLocalMailbox = Get-ADUser -LDAPFilter $ldapHasLocalMailbox -Properties $attributeList
$numLocalMailboxes = $usersWithLocalMailbox | Measure-Object | Select-Object -ExpandProperty Count

# Räkna låsta användare med lokal mailbox
$lockedUsersWithLocalMailbox = Get-ADUser -LDAPFilter $ldapHasLocalMailbox -Properties $attributeList | Where-Object { $_.Enabled -like 'False' }
$numLockedLocalMailboxes = $lockedUsersWithLocalMailbox | Measure-Object | Select-Object -ExpandProperty Count

# Räkna användare med O365-mailbox
$usersWithO365Mailbox = Get-ADUser -LDAPFilter $ldapHasO365Mailbox -Properties $attributeList
$numO365Mailboxes = $usersWithO365Mailbox | Measure-Object | Select-Object -ExpandProperty Count

# Räkna låsta användare med O365-mailbox
$lockedUsersWithO365Mailbox = Get-ADUser -LDAPFilter $ldapHasO365Mailbox -Properties $attributeList | Where-Object { $_.Enabled -like 'False' }
$numLockedO365Mailboxes = $lockedUsersWithO365Mailbox | Measure-Object | Select-Object -ExpandProperty Count

# Skriv sammanställning till fil
$resultstring = "Resultat`nTotalt antal mailboxar: $numTotalMailboxes`nAntal lokala mailboxar: $numLocalMailboxes`n- varav låsta: $numLockedLocalMailboxes`nOffice 365-mailboxar: $numO365Mailboxes`n- varav låsta: $numLockedO365Mailboxes"
$resultstring | Out-File -FilePath $outfile -Encoding utf8

if ( $ListMailboxes ) {
    # Mailboxarna ska listas
    "`n`n=== Användare med mailbox ===" | Out-File -FilePath $outfile -Encoding utf8  -Append
    $usersWithMailbox | Select-Object -Property Name,mail | Out-File -FilePath $outfile -Encoding utf8  -Append

    "`n`n=== Användare med lokal mailbox ===" | Out-File -FilePath $outfile -Encoding utf8  -Append
    $usersWithLocalMailbox | Select-Object -Property Name,mail | Out-File -FilePath $outfile -Encoding utf8  -Append

    "`n`n=== Låsta användare med lokal mailbox ===" | Out-File -FilePath $outfile -Encoding utf8  -Append
    $lockedUsersWithLocalMailbox | Select-Object -Property Name,mail | Out-File -FilePath $outfile -Encoding utf8  -Append

    "`n`n=== Användare med Office 365-mailbox ===" | Out-File -FilePath $outfile -Encoding utf8  -Append
    $usersWithO365Mailbox | Select-Object -Property Name,mail | Out-File -FilePath $outfile -Encoding utf8  -Append

    "`n`n=== Låsta användare med Office 365-mailbox ===" | Out-File -FilePath $outfile -Encoding utf8  -Append
    $lockedUsersWithO365Mailbox | Select-Object -Property Name,mail | Out-File -FilePath $outfile -Encoding utf8  -Append

}