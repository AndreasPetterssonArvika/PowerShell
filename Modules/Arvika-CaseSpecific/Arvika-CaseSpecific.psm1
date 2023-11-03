<#
Modul för funktioner specifika för ett visst ärende
För de fall där det blir mycket kod/många funktioner ska de placeras i en separat modul

Alla funktioner ska ha ett eller flera tillhörande ärendenummer för identifiering
#>

<#
Funktionen tar emot en ADUser från pipeline och lämnar specifika data som en semikolonseparerad textsträng
Ärende #45799
Issue #240
Issue #243, ändrade data i exporten
#>
function Get-SVUserData {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADObject]$ADUser
    )

    begin {}

    process {
        $curUser = $ADUser | Get-ADUser -Properties givenName,SN,mail
        $userdata = $curuser.givenName + ";" + $curUser.SN + ";"  + $curUser.mail
        $ImmutableID = Get-ImmutableIDForUser -ADUser $ADUser
        $userdata += ";" + $ImmutableID
        Write-Output $userdata
        
    }

    end {}
}

<#
Funktioner för att hantera de nya skolanknutna kontona på arvika.se
Ärende #47690
Issue #255
#>

<#
Funktionen skapar en CSV-fil med de aktuella användarna.

Datumen skapas t ex genom ((Get-Date).AddDays(<days>)).Date
VIKTIGT: Ovanstående kommando fungerar
((Get-Date).AddDays(<days>)).DateTime fungerar inte. Det kommandot lämnar ifrån sig en
textsträng och inte ett DateTime-objekt.

Attribut som behövs
- personNummer
- mail
#>
function Get-MIMStartNewUsers {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)][DateTime]$FromTime,
        [Parameter(Mandatory)][DateTime]$ToTime,
        [Parameter(Mandatory)][string]$UserIdentifier,
        [Parameter()][string]$OutputFolder='.'
    )

    $now = Get-Date -Format 'yyyyMMdd_HHmm'

    $outfile = "$OutputFolder\NewMIMUsers_$now.csv"

    "$UserIdentifier,mail" | Out-File -FilePath $outfile -Encoding utf8

    $attrs = @("$UserIdentifier",'mail')

    Get-ADUser -Filter { (whenCreated -ge $FromTime) -and (whenCreated -le $ToTime) } -Properties $attrs | ForEach-Object { $curUserRow="$($_.$UserIdentifier);$($_.mail)";$curUserRow | Out-File -FilePath $outfile -Encoding utf8 -Append }
}

Export-ModuleMember -Function Get-SVUserData
Export-ModuleMember -Function Get-MIMStartNewUsers