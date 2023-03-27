<#
Skriptet hämtar unika gruppnamn från ett Excelblad
Gruppnamnen sparas till en textfil
#>

[cmdletbinding()]
param(
    [Parameter(Mandatory)][string]$Infile,
    [Parameter(Mandatory)][string]$UserColumn
)

Import-Module ImportExcel

$now = get-date -Format 'yyMMdd'
$outfile = "LocalAdmins_$now.txt"

$usernames = Import-Excel -Path $Infile | Select-Object -ExpandProperty $UserColumn

$admUsers =  @{}

foreach ( $username in $usernames) {
    Write-Debug "Kolla: $username"
    $numFound= 0
    $numFound = Get-ADGroup -Filter { name -eq $username } | Measure-Object | Select-Object -ExpandProperty count
    if ( $numFound -gt 0 ) {
        Write-Debug "Grupp: $username"
        $admUsers[$username] = 'group'
    }
}

$admUsers.GetEnumerator() | Sort-Object | Select-Object -ExpandProperty Key | Out-File -FilePath $outfile