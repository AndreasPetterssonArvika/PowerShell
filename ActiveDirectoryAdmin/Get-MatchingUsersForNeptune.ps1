# Skriptet matchar användare från Neptune mot användare i Active Directory
# Indata är en Excelbok som förutsätts ha bladet "all_users"
# Kolumnen som innehåller användanamnet från Neptune förutsätts ha rubriken username
# Skriptet exporterar en ny Excelbok med två blad där det ena innehåller användarnamnen
# som matchar namnstandard för Active Directory

if ( $psISE ) {
    $basePath = Split-Path $psISE.CurrentFile.FullPath
} else {
    $basePath = $PSScriptRoot
}

#$basePath = <manual path>    # Sätt om det ska vara manuell sökväg, lämna utkommenterad annars.
$excelWorkbook = "$basePath\all_users Neptune 210830 Kopia.xlsx"
$excelWorksheet = "all_users"

$neptuneUsers = Import-Excel -Path $excelWorkbook -WorksheetName $excelWorksheet

# Slå upp de som matchar namnstandard och lägg in i excelbladet

$wellFormedUsernameRegex = '^[a-z-]+[.]{1}[a-zA-Z0-9.-]+$'

$attributes = @('mail','SN','sAMAccountName')

$wellFormedUsers = $neptuneUsers | Where-Object { $_.username -match $wellFormedUsernameRegex } | Select-Object -ExpandProperty username

$now = Get-Date -UFormat "%Y%m%d%H%M"

$exportSheetWellFormed = "pWellFormed_$now"

foreach ( $wellFormedUser in $wellFormedUsers ) {
    $curMail = "$wellFormedUser@arvika.se"
    $ldapFilter = "(mail=$wellFormedUser@arvika.se)"
    Get-ADUser -LDAPFilter $ldapFilter -Properties $attributes | Select-Object -Property givenName,SN,mail,sAMAccountName | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetWellFormed -Append
}

# Slå upp de som inte matchar och lägg in i Excelbladet
$badUserNames = $neptuneUsers | Where-Object { $_.username -notmatch $wellFormedUsernameRegex } | Select-Object -ExpandProperty username

$exportSheetBadName = "pBadName_$now"

foreach ( $badUserName in $badUserNames ) {
    $ldapFilter = "(sAMAccountName=$badUserName)"
    Get-ADUser -LDAPFilter $ldapFilter -Properties $attributes | Select-Object -Property givenName,SN,mail,sAMAccountName | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetBadName -Append
}