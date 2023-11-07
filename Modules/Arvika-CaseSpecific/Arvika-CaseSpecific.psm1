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

Genomgående är miljön där användarna med mailboxar skapas den lokala miljön
medan den motsvarande skolmiljö är remote.

Datumen skapas t ex genom ((Get-Date).AddDays(<days>)).Date
VIKTIGT: Ovanstående kommando fungerar
((Get-Date).AddDays(<days>)).DateTime fungerar inte. Det kommandot lämnar ifrån sig en
textsträng och inte ett DateTime-objekt.

Attribut som behövs
- personNummer
- mail
#>
function Get-MIMStartNewUsers {
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][DateTime]$FromDate,
        [Parameter(Mandatory)][DateTime]$ToDate,
        [Parameter(Mandatory)][string]$UserIdentifier,
        [Parameter(Mandatory)][string]$RemoteServer,
        [Parameter(Mandatory)][pscredential]$RemoteCred
    )

    # Attributen som ska hämtas utöver standardattribut
    $attrs = @("$UserIdentifier",'mail')

    # Hashtables för användare
    $localUsers = @{}
    $remoteUsers = @{}

    # Hämtar användare skapade i det definierade tidsintervallet
    # Filtrerar bort användare som saknar personnummer och/eller mailbox
    Get-ADUser -Filter { (whenCreated -ge $FromDate) -and (whenCreated -le $ToDate) -and ($UserIdentifier -like "*" ) -and (mail -like "*") -and (info -like "IDM Managed*") } -Properties $attrs | ForEach-Object { $localUsers[$($_.$UserIdentifier)]=$($_.mail ); $numLocalUsers++ }

    foreach ( $localKey in $localUsers.Keys ) {
        $ldapfilter = "($UserIdentifier=$localKey)"
        Get-ADUser -Server $RemoteServer -Credential $RemoteCred -LDAPFilter $ldapfilter -Properties $attrs | ForEach-Object { $remoteUsers[$($_.$UserIdentifier)]=$($_.mail ) }
    }

    Write-Output "Hittade $($localUsers.Count) nya användare i lokala Active Directory"
    Write-Output "Hittade $($remoteUsers.Count) nya användare i remote Active Directory"
    Write-Output "Hittade $($localUsers.Count - $remoteUsers.Count) nya användare som enbart finns i lokalt Active Directory"

    if ( $($remoteUsers.Count) -gt 0 ) {
        
        # Exportera underlag för migrering
        New-MIMStartMigrationFile -LocalUsers $localUsers -RemoteUsers $remoteUsers

        # Exportera underlag för forwarding
        New-MIMStartForwardingFile -LocalUsers $localUsers -RemoteUsers $remoteUsers
        
        # Dölj mailboxarna
        Set-MIMStartHideUsersFromGAL -LocalUsers $localUsers -RemoteUsers $remoteUsers -WhatIf:$WhatIfPreference

    }

}

<#
Funktionen skapar en underlagsfil för migrering av mailboxar
Filen ska vara en CSV-fil och kolumnen "EmailAddress" är den enda som krävs
Filen hämtar alla nycklar ur remoteUsers och slår upp motsvarande epost-adresser
ur localUsers som underlag för migrering
#>
function New-MIMStartMigrationFile {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][hashtable]$LocalUsers,
        [Parameter(Mandatory)][hashtable]$RemoteUsers,
        [Parameter()][string]$OutputDirectory = '.'
    )
    

    $now = Get-Date -Format 'yyyyMMdd_HHmm'
    
    $outfile = "$OutputDirectory\MIMStartMigrationFile_$now.csv"

    # Behövs rubrikrad?
    'EmailAddress' | Out-File -FilePath $outfile -Encoding utf8

    foreach ( $key in $remoteUsers.Keys ) {
        $localUsers[$key]  | Out-File -FilePath $outfile -Encoding utf8 -Append
    }

}

<#
Funktionen döljer användare från adresslistor
För alla användare i RemoteUsers döljs
motsvarande användare från Localusers från
adressböckerna
#>
function Set-MIMStartHideUsersFromGAL {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][hashtable]$LocalUsers,
        [Parameter(Mandatory)][hashtable]$RemoteUsers
    )

    foreach ( $key in $RemoteUsers.Keys ) {
        $ldapfilter="(mail=$($LocalUsers[$key]))"
        if($PSCmdlet.ShouldProcess($ldapfilter)) {
            Get-ADUser -LDAPFilter $ldapfilter | Set-ADUser -Replace @{msExchHideFromAddressLists=$true}
        }
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

    # Skapa rubrikrad. Kolumnrubrikerna är de som förvänats i Set-MIMStartForwarding
    'MailboxAddress;ForwardToAddress' | Out-File -FilePath $outfile -Encoding utf8

    foreach ( $key in $RemoteUsers.Keys ) {
        "$($LocalUsers[$key]);$($RemoteUsers[$key])" | Out-File -FilePath $outfile -Encoding utf8 -Append
    }

}

<#
Funktionen lägger till användare i grupper för licenstilldelning
och forwarding.
Här fungerar det bra med filen som är underlag för migreringen.
#>
function Add-MIMStartUsersToGroups {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$LicenseGroupName,
        [Parameter(Mandatory)][string]$ForwardGroupName,
        [Parameter(Mandatory)][string]$UserMigrationFile
    )

    $userRows = Import-Csv -Delimiter ';' -Encoding utf8 -Path $UserMigrationFile | Select-Object -ExpandProperty 'EmailAddress'

    foreach ( $user in $userRows ) {
        $ldapfilter="(mail=$user)"
        $curUser = Get-ADUser -LDAPFilter $ldapfilter
        if($PSCmdlet.ShouldProcess($($curUser.Name))) {
            $curUser | Add-ADPrincipalGroupMembership -MemberOf $LicenseGroupName
            $curUser | Add-ADPrincipalGroupMembership -MemberOf $ForwardGroupName
        }
    }
}

<#
Funktionen sätter forwarding för mailboxarna i filen

Funktionen kräver att kopplingen till Exchange Online redan är gjord
Connect-ExchangeOnline på dator som har PSModulen installerad

Filen ska vara semikolonseparerad textfil och måste innehålla två kolumner (det går att ha fler)
Kolumnerna som krävs är:
- MailboxAddress
- ForwardToAddress

För varje användare ska mailadressen finnas
#>
function Set-MIMStartForwarding {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$ForwardingFile
    )

    $forwardingRows = Import-Csv -Delimiter ';' -Encoding utf8 -Path $ForwardingFile

    foreach ( $row in $forwardingRows ) {
        $mailbox=$row.MailboxAddress
        $forwardingAddress = $row.ForwardToAddress
        if ($PSCmdlet.ShouldProcess("Forwarding $mailbox to $forwardingAddress",$mailbox,'forward')) {
            Set-Mailbox -Identity $mailbox -ForwardingSmtpAddress $forwardingAddress -DeliverToMailboxAndForward $false
        }
    }
}


Export-ModuleMember -Function Get-SVUserData
Export-ModuleMember -Function Get-MIMStartNewUsers
Export-ModuleMember -Function Set-MIMStartForwarding
Export-ModuleMember -Function Add-MIMStartUsersToGroups