<#
Skriptet 
#>
[cmdletbinding()]
param(
    [Parameter]
    [string]$LogFile
)

# Platsen för in- och utdata
if ( $psISE ) {
    $BaseFilePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    $BaseFilePath = $PSScriptRoot
}

$ldapfilter = '(&(mail=*@*)(msExchMailboxGuid=''))'


$adusers = Get-ADUser -LDAPFilter $ldapfilter -Properties mail
$adusers | Measure-Object | select-object -ExpandProperty Count

if ( $LogFile ) {
    $LogFilePath = "$BaseFilePath\$LogFile"
    if (!(Test-path -Path $LogFilePath)) {
        New-Item -name $LogFilePath -ItemType File
    }
    $adusers | Select-Object -ExpandProperty mail |  Out-File -FilePath $LogFilePath -Encoding utf8
}