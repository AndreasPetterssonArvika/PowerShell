# Kontrollera en lista med mailadresser mot AD
# Skapa listor för:
# - de som saknar ett värde i msExchMailboxGuid, tolkat som att de inte har någon mailbox alls
# - de som har ett värde i msExchMailboxGuid, tolkat som att de har en mailbox
# - de som har ett värde i msExchMailboxGuid men msexchRemoteRecipientType skild från 4, tolkat som att de har en lokal mailbox
# - de som har ett värde i msExchMailboxGuid och msexchRemoteRecipientType som är 4, tolkat som att de har en Office-365-mailbox

# Platsen för in- och utdata
$BaseFilePath = $PSScriptRoot     # Detta är mappen där skriptet ligger
#$BaseFilePath = 'C:\Temp'        # Alternativ om man vill ha datafilerna någon annanstans



# Indatafil, textfil med en mailadress per rad
$mailAddressListPath = $BaseFilePath + '\inputfile.txt'

#Utdatafiler
$noMailboxFilePath = $BaseFilePath + '\UsersWithoutMailbox.txt'
$hasMailboxFilePath = $BaseFilePath + '\UsersWithMailbox.txt'
$hasLocalMailboxFilePath = $BaseFilePath + '\UsersWithLocalMailbox.txt'
$hasO365FilePath = $BaseFilePath + '\UsersWithOffice365Mailbox.txt'

# Ta bort ev befintliga utdatafiler från tidigare körning
# Rensas de inte fylls de på för varje körning
Remove-Item $noMailboxFilePath
Remove-Item $hasMailboxFilePath
Remove-Item $hasLocalMailboxFilePath
Remove-Item $hasO365FilePath

# Attribut som behöver definieras utöver de som Get-ADUser levererar
$attributeList = @('mail','msExchMailboxGuid','msexchRemoteRecipientType')

# Regex som matchar en tom sträng
$noGuidPattern = '^$'

# Läs in listan med mailadresser
$importedMailAddresses = Get-Content -Path $mailAddressListPath

################################
#                              #
# Slå upp användare till filer #
#                              #
################################

# Slå upp användare utan mailbox
foreach ( $mailAddress in $importedMailAddresses ) {
    Get-ADUser -LDAPFilter "(mail=$mailAddress)" -Properties $attributeList | Where-Object { $_.msExchMailboxGuid -match $noGuidPattern } | Select-Object -ExpandProperty mail | Out-File -FilePath $noMailboxFilePath -Append
}

# Slå upp användare med mailbox
foreach ( $mailAddress in $importedMailAddresses ) {
    Get-ADUser -LDAPFilter "(mail=$mailAddress)" -Properties $attributeList | Where-Object { $_.msExchMailboxGuid -notmatch $noGuidPattern } | Select-Object -ExpandProperty mail | Out-File -FilePath $hasMailboxFilePath -Append
}

# Slå upp användare med lokal mailbox
foreach ( $mailAddress in $importedMailAddresses ) {
    Get-ADUser -LDAPFilter "(mail=$mailAddress)" -Properties $attributeList | Where-Object { ( $_.msExchMailboxGuid -notmatch $noGuidPattern ) -and ( $_.msexchRemoteRecipientType -ne '4' ) } | Select-Object -ExpandProperty mail | Out-File -FilePath $hasLocalMailboxFilePath -Append
}

# Slå upp användare med Office 365-mailbox
foreach ( $mailAddress in $importedMailAddresses ) {
    Get-ADUser -LDAPFilter "(mail=$mailAddress)" -Properties $attributeList | Where-Object { ( $_.msExchMailboxGuid -notmatch $noGuidPattern ) -and ( $_.msexchRemoteRecipientType -eq '4' ) } | Select-Object -ExpandProperty mail | Out-File -FilePath $hasO365FilePath -Append
}

###########################
#                         #
# Räkna användare i filer #
#                         #
###########################

$noMailUsers = Get-Content $noMailboxFilePath
$numNoMailUsers = $noMailUsers | measure | Select-Object -ExpandProperty count

$mailUsers = Get-Content $hasMailboxFilePath
$numMailUsers = $mailUsers | measure | Select-Object -ExpandProperty count

$localMailUsers = Get-Content $hasLocalMailboxFilePath
$numLocalUsers = $localMailUsers | measure | Select-Object -ExpandProperty count

$O365MailUsers = Get-Content $hasO365FilePath
$numO365Users = $O365MailUsers | measure | Select-Object -ExpandProperty count

Write-Host "`n`nAntal användare utan mailbox: $numNoMailUsers"
Write-Host "Antal användare med mailbox: $numMailUsers"
Write-Host "Antal användare med lokal mailbox: $numLocalUsers"
Write-Host "Antal användare med Office 365-mailbox:  $numO365Users"