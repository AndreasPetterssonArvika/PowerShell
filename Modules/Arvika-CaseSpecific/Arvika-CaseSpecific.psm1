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
Funktionen hämtar de aktuella användarna från det lokala directoryt
och motsvarande användare från ett remote-directory.

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
        [Parameter(Mandatory)][DateTime]$FromDate,
        [Parameter(Mandatory)][DateTime]$ToDate,
        [Parameter(Mandatory)][string]$UserIdentifier,
        [Parameter(Mandatory)][string]$RemoteServer,
        [Parameter(Mandatory)][pscredential]$RemoteCred
    )

    

    $attrs = @("$UserIdentifier",'mail')

    $localUsers = @{}
    $remoteUsers = @{}

    # Filtrerar bort användare som saknar personnummer och/eller mailbox
    Get-ADUser -Filter { (whenCreated -ge $FromDate) -and (whenCreated -le $ToDate) -and ($UserIdentifier -like "*" ) -and (mail -like "*") } -Properties $attrs | ForEach-Object { $localUsers[$($_.$UserIdentifier)]=$($_.mail ); $numLocalUsers++ }

    foreach ( $localKey in $localUsers.Keys ) {
        $ldapfilter = "($UserIdentifier=$localKey)"
        Get-ADUser -Server $RemoteServer -Credential $RemoteCred -LDAPFilter $ldapfilter -Properties $attrs | ForEach-Object { $remoteUsers[$($_.$UserIdentifier)]=$($_.mail ) }
    }

    <#
    $numLocalUsers = $localUsers.Count
    $numRemoteUsers = $remoteUsers.Count
    Write-Verbose "Found $numLocalUsers users in local Active Directory"
    Write-Verbose "Found $numRemoteUsers users in remote Active Directory"
    #>

    Write-Verbose "Found $($localUsers.Count) users in local Active Directory"
    Write-Verbose "Found $($remoteUsers.Count) users in remote Active Directory"

    if ( $($remoteUsers.Count) -gt 0 ) {
        
        # Exportera underlag för migrering
        New-MIMStartMigrationFile -UsersToMigrate $remoteUsers

        # Exportera underlag för forwarding
        New-MIMStartForwardingFile -LocalUsers $localUsers -RemoteUsers $remoteUsers
        
        # Dölj mailboxarna
        Set-MIMStartHideUsersFromGAL -LocalUsers $localUsers -RemoteUsers $remoteUsers

    }

}

<#
Funktionen skapar en underlagsfil för migrering av mailboxar
#>
function New-MIMStartMigrationFile {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][hashtable]$UsersToMigrate,
        [Parameter()][string]$OutputDirectory = '.'
    )
    

    $now = Get-Date -Format 'yyyyMMdd_HHmm'
    
    $outfile = "$OutputDirectory\MIMStartMigrationFile_$now.csv"

    # Behövs rubrikrad?
    #'mail' | Out-File -FilePath $outfile -Encoding utf8

    foreach ( $user in $UsersToMigrate ) {
        $($user.Values)  | Out-File -FilePath $outfile -Encoding utf8 -Append
    }

}

<#
Funktionen döljer användare från adresslistor
För alla användare i RemoteUsers döljs
motsvarande användare från Localusers från
adressböckerna
#>
function Set-MIMStartHideUsersFromGAL {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][hashtable]$LocalUsers,
        [Parameter(Mandatory)][hashtable]$RemoteUsers
    )

    foreach ( $key in $RemoteUsers.Keys ) {
        $ldapfilter="(mail=$($LocalUsers[$key]))"
        Get-ADUser -LDAPFilter $ldapfilter | Set-ADUser -Replace @{msExchHideFromAddressLists=$true} -WhatIf
    }

}

<#
Funktionen skapar ett underlag för vidarebefordran från en Office 365-mailbox
till en annan mail-adress
#>
function New-MIMStartForwardingFile {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][hashtable]$LocalUsers,
        [Parameter(Mandatory)][hashtable]$RemoteUsers,
        [Parameter()][string]$OutputDirectory = '.'
    )

    $now = Get-Date -Format 'yyyyMMdd_HHmm'

    $outfile = "$OutputDirectory\MIMStartForwardingFile_$now.csv"

    # Skapa rubrikrad
    'O365Mail;ForwardToMail' | Out-File -FilePath $outfile -Encoding utf8

    foreach ( $key in $RemoteUsers.Keys ) {
        "$($LocalUsers[$key]);$($RemoteUsers[$key])" | Out-File -FilePath $outfile -Encoding utf8 -Append
    }

}


Export-ModuleMember -Function Get-SVUserData
Export-ModuleMember -Function Get-MIMStartNewUsers