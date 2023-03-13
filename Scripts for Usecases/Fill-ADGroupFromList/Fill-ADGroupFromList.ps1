<#
Skriptet tar en excelfil med indata
En av kolumnerna måste innehålla data som förväntas finnas i Active Directory och
ha samma kolumnrubrik som attributet
Namnet på kolumnrubeiken/attributet anges i parametern IDField
Namnet på gruppen där personerna ska läggas in ska anges i parametern ADGroup

Om det finns personer som inte hittas och läggs in i funktionen rapporteras
antalet personer och deras ID läggs i en Excelfil.
Denna fil kan om man förser kolumnen med ID användas som indata i det här skriptet.
#>

[cmdletbinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string]$UserlistFile,    # Fil med personerna som ska hittas i AD
    [Parameter(Mandatory)][string]$ADGroup,         # Namnet på gruppen där de ska läggas till
    [Parameter(Mandatory)][string]$IDField,         # Namnet på kolumnen med sökdatat
    [Parameter(Mandatory)][switch]$ID13             # Anger att konvertering från id med 13 tecken måste göras
)

Import-Module ImportExcel

function ConvertTo-IDKey12 {
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName = 'IDK13')]
        [string]$IDKey13,
        [Parameter(ParameterSetName = 'IDK11')]
        [string]$IDKey11,
        [Parameter(ParameterSetName = 'IDK10')]
        [string]$IDKey10
    )

    $tKey = ''

    if ( $PSCmdlet.ParameterSetName -eq 'IDK13') {

        # Konvertera från 13 till 12 tecken
        Write-Debug "Converting $IDKey13"
        $yyyymmdd=$IDKey13.Substring(0,8)
        $nums=$IDKey13.Substring(9,4)
        $tKey="$yyyymmdd$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK11') {

        # Konvertera från 11 till 12 tecken
        Write-Debug "Converting $IDKey11"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey11.Substring(0,2))
        $mmdd=$IDKey11.Substring(2,4)
        $nums=$IDKey11.Substring(7,4)
        $tKey="$year$mmdd$nums"

    } elseif ( $PSCmdlet.ParameterSetName -eq 'IDK10') {

        # Konvertera från 10 till 12 tecken
        Write-Debug "Converting $IDKey10"
        $year=(Get-Culture).Calendar.ToFourDigitYear($IDKey10.Substring(0,2))
        $mmdd=$IDKey10.Substring(2,4)
        $nums=$IDKey10.Substring(6,4)
        $tKey="$year$mmdd$nums"

    } else {
        # Okänt parmeterset
        Write-Error "Unknown Parameterset"
    }

    return $tKey

}

$users = Import-Excel -Path $UserlistFile

$usersNotFound = @{}

foreach ( $user in $users ) {
    $curID = $user.$IDField
    if ( $ID13 ) {
        $curID = ConvertTo-IDKey12 -IDKey13 $curID
        write-debug "Konverterat ID13 -> ID12"
    }
    $ldapfilter = "($IDField=$curID)"
    $curUser = Get-ADUser -LDAPFilter $ldapfilter

    if ( $curUser ) {
        # Användaren hittad, lägg till i grupp
        if ( $PSCmdlet.ShouldProcess($curUser)) {
            $curUser | Add-ADPrincipalGroupMembership -MemberOf $ADGroup
        }
    } else {
            $usersNotFound[$curID] = 'notfound'
    }
}

$numMissingUsers = $usersNotFound.Count

if ( $numMissingUsers -gt 0 ) {
    Write-Host "Antal användare som inte hittades: $numMissingUsers"
    # Exportera ID för dessa användare
    $now = Get-Date -Format 'yymmdd_HHmm'
    $missingUserFile = "MissingUsers_$now.xlsx"
    $usersNotFound.Keys | Export-Excel -Path $missingUserFile -WorksheetName 'Missing'
}