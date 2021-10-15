# Skriptet matchar användare från Neptune mot användare i Active Directory
# Indata är en Excelbok som förutsätts ha bladet "all_users"
# Kolumnen som innehåller användanamnet från Neptune förutsätts ha rubriken username
# Skriptet jämför de uppslagna användarnamnen mot formen på ett välformat användarnamn,
# slår upp användarna och kontrollerar om deras userPrincipalname matchar attributet mail.
# Alla hittade användare skrivs till nya blad (upp till fyra) i Excelboken.

if ( $psISE ) {
    $basePath = Split-Path $psISE.CurrentFile.FullPath
} else {
    $basePath = $PSScriptRoot
}

#$basePath = '<manuell sökväg>'    # Sätt om det ska vara manuell sökväg, lämna utkommenterad annars.
$excelWorkbook = "$basePath\<Excelboksnamn>"
$excelWorksheet = "all_users"

$neptuneUsers = Import-Excel -Path $excelWorkbook -WorksheetName $excelWorksheet

$notFoundInActiveDirectory = @{}
$notFoundKey = 'NotFound'

# Slå upp de som matchar namnstandard och lägg in i excelbladet

$wellFormedUsernameRegex = '^[a-z-]+[.]{1}[a-zA-Z0-9.-]+$'

$attributes = @('mail','SN','sAMAccountName')

$wellFormedUsers = $neptuneUsers | Where-Object { $_.username -match $wellFormedUsernameRegex } | Select-Object -ExpandProperty username

$now = Get-Date -UFormat "%Y%m%d%H%M"

$exportSheetWellFormedMatched = "WellFormedMatched_$now"
$exportSheetWellFormedMisMatch = "WellFormedMisMatch_$now"

foreach ( $wellFormedUser in $wellFormedUsers ) {
    $curMail = "$wellFormedUser@arvika.se"
    $ldapFilter = "(userPrincipalname=$wellFormedUser@arvika.se)"
    $curUser = Get-ADUser -LDAPFilter $ldapFilter -Properties $attributes
    $numUser = $curUser | measure | Select-Object -ExpandProperty count
    if ( $numUser -gt 0 ) {
        if ( $curUser.userPrincipalname -eq $curUser.mail ) {
            # UPN och mail matchar
            $curUser | Select-Object -Property givenName,SN,userPrincipalName,mail | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetWellFormedMatched -Append
        } else {
            # UPN och mail matchar inte
            $curUser | Select-Object -Property givenName,SN,userPrincipalName,mail | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetWellFormedMisMatch -Append
        }
    } else {
        # Lägg till ej hittat användarnamn
        $notFoundInActiveDirectory.Add($wellFormedUser,$notFoundKey)
    }
}

# Slå upp de som inte matchar och lägg in i Excelbladet
$badUserNames = $neptuneUsers | Where-Object { $_.username -notmatch $wellFormedUsernameRegex } | Select-Object -ExpandProperty username

$exportSheetBadNameMatched = "BadNameMatched_$now"
$exportSheetBadNameMisMatch = "BadNameMisMatch_$now"

foreach ( $badUserName in $badUserNames ) {
    $ldapFilter = "(sAMAccountName=$badUserName)"
    #Get-ADUser -LDAPFilter $ldapFilter -Properties $attributes | Select-Object -Property givenName,SN,userPrincipalName,mail | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetBadName -Append
    $curUser = Get-ADUser -LDAPFilter $ldapFilter -Properties $attributes
    $numUser = $curUser | measure | Select-Object -ExpandProperty count
    if ( $numUser -gt 0 ) {
        if ( $curUser.userPrincipalname -eq $curUser.mail ) {
            # UPN och mail matchar
            $curUser | Select-Object -Property givenName,SN,userPrincipalName,mail | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetBadNameMatched -Append
        } else {
            # UPN och mail matchar inte
            $curUser | Select-Object -Property givenName,SN,userPrincipalName,mail | Export-Excel -Path $excelWorkbook -WorksheetName $exportSheetBadNameMisMatch -Append
        }
    } else {
        # Lägg till ej hittat användarnamn
        $notFoundInActiveDirectory.Add($badUserName,$notFoundKey)
    }

}

# Exportera ej hittade Neptune-användare till Excelbladet
$notFoundWorkSheet = "NotFoundInAD_$now"
foreach ( $username in $notFoundInActiveDirectory.Keys ) {
    $username | Export-Excel -Path $excelWorkbook -WorksheetName $notFoundWorkSheet -Append
}
