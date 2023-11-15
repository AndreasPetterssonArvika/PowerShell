<#
Funktioner för att arbeta med Edlevos API:er
#>

function Get-EdlevoPerson {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)][string]$BaseDomain,
        [Parameter(Mandatory)][string]$PersonID,
        [Parameter(Mandatory)][string]$LicenseKey
    )

    $statusSuccess = '200'

    $uri="https://$basedomain/WE.Education.Integration.Host.Proxy/LES/Person/v4/Person/GetPersonInfo?LicenseKey=$LicenseKey&PersonID=$PersonID"

    [string]$response = Invoke-RestMethod -Uri $Uri -StatusCodeVariable 'scv' -Method Get

    if ( $scv -match $statusSuccess ) {
        # Fått tillbaka en person

        # Konvertera svaret till bytes för att klippa av inledande BOM (6 bytes)
        $responseBytes = [System.Text.Encoding]::UTF8.GetBytes($response)
        
        $responseBytesLength = $responseBytes.Length - 1
        $xmlBytes = $responseBytes[6..$responseBytesLength]

        # Konvertera bytes till XML-data
        [xml]$xmlContent = [System.Text.Encoding]::UTF8.GetString($xmlBytes)

        return $xmlContent

    } else {

        return $null
    }

}

function Get-EdlevoPersonFromFile {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$XmlInput
    )

    [xml]$EdlevoPerson=Get-Content -Path $XmlInput

    Write-Output $EdlevoPerson

    $boardgameItems = $EdlevoPerson | Select-Xml -XPath "//userid"
}

function Get-EdlevoOrganization {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)][string]$BaseDomain,
        [Parameter()][string]$UnitID='NoID',
        [Parameter(Mandatory)][string]$LicenseKey
    )

    $statusSuccess = '200'

    $uri="https://$basedomain/WE.Education.Integration.Host.Proxy/LES/Organization/v10/Organization/GetUpperSecondarySchoolOrganization?LicenseKey=$LicenseKey"

    if ( $UnitID -match 'NoID' ) {
        # Ingen enskild enhet
        Write-Information "Ingen särskild enhet"
    } else {
        # En enhet ska slås upp, komplettera
        $uri = $uri + "&UnitID=$UnitID"
    }

    [string]$response = Invoke-RestMethod -Uri $Uri -StatusCodeVariable 'scv' -Method Get

    if ( $scv -match $statusSuccess ) {
        # Fått tillbaka en data

        # Konvertera svaret till bytes för att klippa av inledande BOM (6 bytes)
        $responseBytes = [System.Text.Encoding]::UTF8.GetBytes($response)
        
        $responseBytesLength = $responseBytes.Length - 1
        $xmlBytes = $responseBytes[6..$responseBytesLength]

        # Konvertera bytes till XML-data
        [xml]$xmlContent = [System.Text.Encoding]::UTF8.GetString($xmlBytes)

        return $xmlContent

    } else {

        return $null
    }

}

function Get-EdlevoOrganizationStaff {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)][xml]$OrgDataXml
    )

    # TODO:Kolla sökningen
    $orgStaff =  $OrgDataXml.SelectNodes("//person[institutionrole/@institutionroletype='Staff']")

    <#
    
    foreach ( $node in $orgStaff ) {

    }

    return $orgStaff
    #>
}

<#
Funktionen hämtar uppgifter från Active Directory och 
uppdaterar motsvarande epost-adresser i Edlevo.

Uppgifterna hämtas från AD baserat på en eller båda av följande
- SearchBase
- LDAPFilter

Funktionen kan hämta data från ett remote directory om man anger lämpliga parametrar

Funktionen begränsar uppslaget till att enbart beröra de senast ändrade användarna
Om parametern DaysSinceUserChange sätts till 0 begränsas inte uppslaget i tid, utan alla användare berörs
#>
function Update-EdlevoEmailFromActiveDirectory {
    [cmdletbinding(SupportsShouldProcess,DefaultParameterSetName='LocalDirectory')]
    param (
        [Parameter(ParameterSetName='RemoteDirectory',Mandatory=$False)][switch]$RemoteDirectory=$False,
        [Parameter(ParameterSetName='RemoteDirectory',Mandatory=$True)][string]$RemoteServer,
        [Parameter(ParameterSetName='RemoteDirectory',Mandatory=$True)][pscredential]$RemoteCredential,
        #[Parameter(Mandatory)][string]$EdlevoAPILicenseKey,
        [Parameter(Mandatory)][string]$EdlevoAPIURI,
        [Parameter()][string]$SearchBase,
        [Parameter()][string]$LDAPFilter,
        [Parameter(Mandatory=$True)][string]$UserIdentifier,
        [Parameter(Mandatory=$True)][string]$MailAttribute,
        [Parameter()][Int32]$DaysSinceUserChange=1
    )

    # Kontrollera att sökbegrepp angivits
    if ( (-not $SearchBase) -and ( -not $LDAPFilter) ) {
        # Inget sökbegrepp angivet
        throw 'Inga sökbegrepp angivna. Ange minst ett av SearchBase eller LDAPFilter'
    }

    if ( $DaysSinceUserChange -gt 0 ) {
        $ChangeTimeCutoff = (Get-Date).AddDays(-$DaysSinceUserChange).ToString('yyyyMMddHHmmss.0Z')
        $LDAPFilter = "(&$LDAPFilter(whenChanged>=$ChangeTimeCutoff))"
    }

    # Hämta data från Active Directory
    Write-Debug 'Hämtar data från Active Directory'
    Write-Debug "SearchBase: $SearchBase"
    Write-Debug "LDAP-filter: $LDAPFilter"
    if ( $RemoteDirectory ) {
        [hashtable]$userData = Get-EdlevoMailUpdateFromActiveDirectory -RemoteDirectory -RemoteServer $RemoteServer -RemoteCredential $RemoteCredential -SearchBase $SearchBase -LDAPFilter $LDAPFilter -UserIdentifier $UserIdentifier -MailAttribute $MailAttribute
    } else {
        [hashtable]$userData = Get-EdlevoMailUpdateFromActiveDirectory -SearchBase $SearchBase -LDAPFilter $LDAPFilter -UserIdentifier $UserIdentifier -MailAttribute $MailAttribute
    }
    
    # Meddela hur många som hittades
    if ( $DebugPreference ) {
        $numUsers = $userData.Keys | Measure-Object | Select-Object -ExpandProperty Count
        Write-Debug "Hittade $numUsers ändrade de senaste $DaysSinceUserChange dagarna"
    }

    # Loopa igenom alla användare och uppdatera epost-adressen
    foreach ( $key in $userdata.Keys ) {
        [xml]$curEmailXML = New-EdlevoPersonEmailXML -UserIdentifier $key -mailAddress $($userData[$key]) -Verbose
        Write-Debug $curEmailXML
        if ( $PSCmdlet.ShouldProcess($key) ) {
            $status = Send-PersonEmailXML -EdlevoUri $EdlevoAPIURI -PersonXML $curEmailXML
            if ( $status -eq '200' ) {
                # Gick bra, gör inget
            } else {
                # Statuskoden säger att det inte gick bra. Meddela.
                Write-Output "Skapa bättre feedback här!"
                Write-Output "Fel för $key $($userData[$key])"
            }
        }
    }
}

<#
Funktionen slår upp data ur ett Active Directory och returnerar en hashtable med
användaridentifieraren som nyckel och epost-adressen som värde
#>
function Get-EdlevoMailUpdateFromActiveDirectory {
    [cmdletbinding(DefaultParameterSetName='LocalDirectory')]
    param (
        [Parameter(ParameterSetName='RemoteDirectory',Mandatory=$False)][switch]$RemoteDirectory=$False,
        [Parameter(ParameterSetName='RemoteDirectory',Mandatory=$True)][string]$RemoteServer,
        [Parameter(ParameterSetName='RemoteDirectory',Mandatory=$True)][pscredential]$RemoteCredential,
        [Parameter()][string]$SearchBase,
        [Parameter()][string]$LDAPFilter,
        [Parameter(Mandatory=$True)][string]$UserIdentifier,
        [Parameter(Mandatory=$True)][string]$MailAttribute
    )

    $UserProps = @($UserIdentifier,$MailAttribute)
    Write-Debug 'Användarattribut'
    $UserProps | ForEach-Object { Write-Debug $_ }

    $UserData = @{}

    if ( $RemoteDirectory ) {
        Write-Debug "Hämtar från remote AD"
        Get-ADUser -LDAPFilter $Ldapfilter -Properties $UserProps -Server $RemoteServer -Credential $RemoteCredential | ForEach-Object { $UserData[$_.$UserIdentifier] = $_.$MailAttribute }
    } else {
        Write-Debug "Hämtar från lokalt AD"
        Get-ADUser -LDAPFilter $Ldapfilter -Properties $UserProps | ForEach-Object { $UserData[$_.$UserIdentifier] = $_.$MailAttribute }
    }

    return $UserData
}

<#
Funktionen skapar XML för att skriva en epostadress till Edlevo
#>
function New-EdlevoPersonEmailXML {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$UserIdentifier,
        [Parameter(Mandatory)][string]$mailAddress
    )

    $userIDElement='<userid useridtype="PID">' + $UserIdentifier + '</userid>'
    $emailElement='<emailworkschool>' + $mailAddress + '</emailworkschool>'
    Write-Debug $userIDElement
    Write-Debug $emailElement

    $xmlStart='<?xml version="1.0" encoding="utf-8"?>
    <person-root 
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
        xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
        xmlns="http://open.tieto.com/edu/person/v4">
    <properties>
        <datetime>2023-10-27T11:10:00</datetime>
        <datasource>Tieto Education</datasource>
        <type>PersonInfo</type>
    </properties>
    <person>'

    $xmlEnd='</person>
    </person-root>'

    $newPersonXML = $xmlStart + $userIDElement + $emailElement + $xmlEnd
    
    return $newPersonXML
    
}

<#
Funktionen skapar en URI mot Edlevos API:er
#>
function New-EdlevoURI {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$True)][string]$BaseDomain,
        [Parameter(Mandatory=$True)][string]$APIEndpoint,
        [Parameter(Mandatory=$True)][string]$LicenseKey
    )

    #$basedomain='eduint.arvika.se'

    #$apiEndpoint='Person/v4/Person/UpdateContactInformation'

    $uri = 'https://' + $BaseDomain + '/WE.Education.Integration.Host.Proxy/LES/' + $APIEndpoint + '?LicenseKey=' + $LicenseKey

    return $uri

}

function Send-PersonEmailXML {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$EdlevoUri,
        [Parameter(Mandatory)][xml]$PersonXML
    )

    $statusSuccess = '200'

    $contentType='text/plain'

    #<#
    $response = Invoke-RestMethod -Uri $EdlevoUri -Method Post -StatusCodeVariable 'scv' -Body $PersonXML -ContentType $contentType

    Write-Debug "Status code: $scv"

    if ( $scv -match $statusSuccess ) {
        # Statuskod för OK
        return $scv
    } else {
        return $scv
    }
}

Export-ModuleMember -Function Get-EdlevoPerson
Export-ModuleMember -Function Get-EdlevoPersonFromFile
Export-ModuleMember -Function Get-EdlevoOrganization
Export-ModuleMember -Function Get-EdlevoOrganizationStaff
Export-ModuleMember -Function Update-EdlevoEmailFromActiveDirectory
Export-ModuleMember -Function Get-EdlevoMailUpdateFromActiveDirectory
Export-ModuleMember -Function New-EdlevoURI