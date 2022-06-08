<#
Skriptet slår upp alla användare i Active Directory som har ett mailattribut men inte har
mailbox GUID

Anger man ett värde för OutFile skrivs aktuella epost-adresser till filen
#>
[cmdletbinding()]
param(
    [Parameter()]
    [string]$OutFile
)

# Sökvägen till skriptet
if ( $psISE ) {
    $BaseFilePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    $BaseFilePath = $PSScriptRoot
}

# Attribut som ska slås upp ur AD
$attributeList = @('mail','msExchMailboxGuid')

# Regex som matchar en tom sträng
$noGuidPattern = '^$'

# Hämta lokalt domännamn
$localDomain = ($env:USERDNSDOMAIN).ToLower()

# LDAP-filter för användare med mailattributet satt til något som liknar en mail-adress
$ldapfilter = "(mail=*@$localDomain)"
Write-Debug $ldapfilter

# Slå upp alla användare med mailadress i attributet mail, sålla fram de som inte har msExchMailboxGuid
$adusers = Get-ADUser -LDAPFilter $ldapfilter -Properties $attributeList | Where-Object { $_.msExchMailboxGuid -match $noGuidPattern}

# Räkna antalet användare
$numWOMailbox = $adusers | Measure-Object | select-object -ExpandProperty Count
Write-Host "Det finns $numWOMailbox användare som har mailattributet satt, men ingen mailbox"

# Om det finns ett angivet filnamn, skriv användare till filen
if ( $OutFile ) {
    $OutFilePath = "$BaseFilePath\$OutFile"
    Write-Debug $OutFilePath
    if (!(Test-path -Path $OutFilePath)) {
        New-Item -name $OutFile -ItemType File -Path $BaseFilePath
    }
    Write-Verbose "Loggfil angiven, exporterar alla epost-adresser till $OutFilePath"
    $adusers | Select-Object -ExpandProperty mail |  Out-File -FilePath $OutFilePath -Encoding utf8
}