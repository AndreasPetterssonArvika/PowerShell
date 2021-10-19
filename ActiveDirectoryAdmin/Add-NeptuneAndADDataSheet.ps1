# Skriptet tar filen med Neptune-användare och relaterade epost-adresser och
# skapar ett nytt blad i arbetsboken med tre kolumner
# - NeptuneUser, användarnamnet i Neptune
# - mail, epost-adress
# - userPrincipalname, användarens inloggningsnamn i Active Directory


Function Get-FileName($initialDirectory)
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “Excelfiler (*.xlsx)| *.xlsx”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

if ( $psISE ) {
    $basePath = Split-Path $psISE.CurrentFile.FullPath
} else {
    $basePath = $PSScriptRoot
}


$excelWorkbook = Get-FileName($basePath)
$excelWorksheet = "all_users"

$neptuneUsers = Import-Excel -Path $excelWorkbook -WorksheetName $excelWorksheet

$userObjectArray = @()

foreach ( $row in $neptuneUsers ) {
    $newUser = @([PSCustomObject]@{ NeptuneUser = $row.username ; mail = $row.mail ; userPrincipalName = '' })

    $userObjectArray = $userObjectArray + $newUser
}

foreach ( $userObject in $userObjectArray ) {
    $tempMail = $userObject.mail
    $ldapFilter = "(mail=$tempMail)"
    $userObject.userPrincipalName = Get-ADUser -LDAPFilter $ldapFilter -Properties userPrincipalName | Select-Object -ExpandProperty userPrincipalname
}

$now = Get-Date -UFormat "%Y%m%d%H%M"

$outputWorksheet = "MatchData_$now"

$userObjectArray | Export-Excel -Path $excelWorkbook -WorksheetName $outputWorksheet -Append