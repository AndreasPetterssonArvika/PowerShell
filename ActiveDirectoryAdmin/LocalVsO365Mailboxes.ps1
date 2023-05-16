# Kontrollera mailboxar
# Kategorier av intresse
# - de som har ett värde i msExchMailboxGuid, tolkat som att de har en mailbox
# - de som har ett värde i msExchMailboxGuid men msexchRemoteRecipientType skild från 4, tolkat som att de har en lokal mailbox
# - de som har ett värde i msExchMailboxGuid och msexchRemoteRecipientType som är 4, tolkat som att de har en Office-365-mailbox
#
# Skriptet räknar också de användare som är låsta, eftersom de i de flesta fall är funktionskonton som inte loggar in utan
# används av andra användare.

$ldapHasMailbox = '(msExchMailboxGuid=*)'
$ldapHasLocalMailbox = '(&(msExchMailboxGuid=*)(!(msexchRemoteRecipientType=4)))'
$ldapHasO365Mailbox = '(&(msExchMailboxGuid=*)(msexchRemoteRecipientType=4))'

# Attribut som behöver definieras utöver de som Get-ADUser levererar
$attributeList = @('mail','msExchMailboxGuid','msexchRemoteRecipientType','Enabled')

# Räkna användare med mailbox
$numTotalMailboxes = Get-ADUser -LDAPFilter $ldapHasMailbox -Properties $attributeList | Measure-Object | Select-Object -ExpandProperty Count

# Räkna användare med lokal mailbox
$numLocalMailboxes = Get-ADUser -LDAPFilter $ldapHasLocalMailbox -Properties $attributeList | Measure-Object | Select-Object -ExpandProperty Count

# Räkna låsta användare med lokal mailbox
$numLockedLocalMailboxes = Get-ADUser -LDAPFilter $ldapHasLocalMailbox -Properties $attributeList | Where-Object { $_.Enabled -like 'False' } | Measure-Object | Select-Object -ExpandProperty Count

# Räkna användare med O365-mailbox
$numO365Mailboxes = Get-ADUser -LDAPFilter $ldapHasO365Mailbox -Properties $attributeList | Measure-Object | Select-Object -ExpandProperty Count

# Räkna låsta användare med O365-mailbox
$numLockedO365Mailboxes = Get-ADUser -LDAPFilter $ldapHasO365Mailbox -Properties $attributeList | Where-Object { $_.Enabled -like 'False' } | Measure-Object | Select-Object -ExpandProperty Count

$resultstring = "`n`nResultat`nTotalt antal mailboxar: $numTotalMailboxes`nAntal lokala mailboxar: $numLocalMailboxes`n- varav låsta: $numLockedLocalMailboxes`nOffice 365-mailboxar: $numO365Mailboxes`n- varav låsta: $numLockedO365Mailboxes"

$resultstring | Out-File -FilePath .\LocalVsO365Mailboxes.txt