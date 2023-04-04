<#
Skriptet uppdaterar ett excelblad baserat på data från AD
#>

#$ExcelWorkbook = 'C:\temp\DF_LS41736\Labb.xlsx'
$ExcelWorkbook = 'C:\temp\DF_LS41736\Användarunderlag DF Respons.xlsx'
$ExcelSheet = 'Blad1'
$ADAttribs = @('sAMAccountname','givenName','SN')
$outfile = 'C:\temp\DF_LS41736\DF-list.csv'
$missingfile= 'C:\temp\DF_LS41736\DF_missing.txt'

$WB = Import-Excel -Path $ExcelWorkbook

$WB | Select-Object -First 5 -ExpandProperty 'E-post adress'

$mailPattern="@arvika.se$"
$textPattern="[\w]{1,}"
$delim=';'

$output = "E-post adress;Enhet;Yrkestitel;SAM AccountName;Förnamn;Efternamn"
$output | Out-File -FilePath $outfile -Encoding utf8

#<#
foreach ( $row in $WB ) {
    $curMail=$row.'E-post adress'
    if ( $curMail -match $mailPattern ) {
        $ldapfilter = "(mail=$curMail)"
        $curAttribs = Get-ADUser -LDAPFilter $ldapfilter -Properties $ADAttribs
        if ( $curAttribs.sAMAccountname -match $textPattern ) {
            $output = $curMail + $delim
            $output += $row.'Enhet' + $delim
            $output += $row.'Yrkestitel' + $delim
            $output += $curAttribs.sAMAccountname + $delim
            $output += $curAttribs.givenName + $delim
            $output += $curAttribs.SN
            $output | Out-File -FilePath $outfile -Encoding utf8 -Append
        } else {
            $curMail | Out-File -FilePath $missingfile -Encoding utf8 -Append
        }
        
    } 
    
}

#>