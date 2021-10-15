# Kontrollera en mailadress mot AD
#
# Förutsättningar
# - de som saknar ett värde i msExchMailboxGuid, tolkas som att de inte har någon mailbox alls
# - de som har ett värde i msExchMailboxGuid men msexchRemoteRecipientType skild från 4, tolkas som att de har en lokal mailbox
# - de som har ett värde i msExchMailboxGuid och msexchRemoteRecipientType som är 4, tolkas som att de har en Office-365-mailbox


# Attribut som behöver definieras utöver de som Get-ADUser levererar
$attributeList = @('mail','msExchMailboxGuid','msexchRemoteRecipientType')

# Regex som matchar en tom sträng
$noGuidPattern = '^$'

$mailAddress = '<mailaddress>'

# Slå upp användare med attribut
$curuser = Get-ADUser -LDAPFilter "(mail=$mailAddress)" -Properties $attributeList

if ( $curUser.msExchMailboxGuid -match $noGuidPattern ) {
    # Ingen maiboxGuid, alltså ingen mailbox
    $mailboxResult = 'ingen'
} elseif ( $curuser.msexchRemoteRecipientType -eq 4 ) {
    # mailboxGuid och msexchRemoteRecipientType lika med 4, alltså Office 365
    $mailboxResult = 'Office 365'
} else {
    # mailboxGuid och msexchRemoteRecipientType inte lika med 4, alltså lokal Exchange
    $mailboxResult = 'lokal Exchange'
}

Write-Host "Användaren har $mailboxResult mailbox"