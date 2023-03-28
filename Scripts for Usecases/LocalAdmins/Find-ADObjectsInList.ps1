<#
Skriptet läser in en 
#>

[cmdletbinding()]
param (
    [Parameter(Mandatory)][string]$InputFile,
    [Parameter(Mandatory)][string]$FileHeader,
    [Parameter(Mandatory)][string]$OU
)

<#
$InputFile = 'C:\temp\LAPSComputers.csv'
$FileHeader = 'AssetName'
$OU='OU=KLS Biblioteket Datorer,OU=KLS Biblioteket,OU=KLS Kultur och fritid,OU=KLS Kommunledningsstab,DC=arvika,DC=se'
#>

[hashtable]$ImportedNames = @{}

# Hämta alla unika namn från indatafilen till en hashtable
Import-Csv -Path $InputFile -Delimiter ';' | select-object -ExpandProperty $FileHeader | ForEach-Object { $ImportedNames[$_]='ADObject' }

#$ImportedNames.Keys


# Slå upp alla datorer ur OU och kolla mot hashtable med namn
$objNotInList = @{}
#Get-ADObject -Filter * -SearchBase $OU -SearchScope Subtree
Get-ADObject -Filter * -SearchBase $OU -SearchScope Subtree -Properties cn | Select-Object -ExpandProperty cn | ForEach-Object { if ( $ImportedNames.ContainsKey($_) ) { $objNotInList[$_]='found' } else { $objNotInList[$_]='missing' } }

foreach ( $key in $objNotInList.Keys ) {
    $status = $objNotInList[$key]
    Write-output "$key`: $status"
}