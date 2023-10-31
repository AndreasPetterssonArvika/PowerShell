<#
Funktioner för att hantera Edlevo API:er
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
    $orgStaff = $OrgDataXml.SelectNodes("//person[institutionrole/@institutionroletype='Staff']")

    foreach ( $node in $orgStaff ) {

    }

    return $orgStaff
}

Export-ModuleMember Get-EdlevoPerson
Export-ModuleMember Get-EdlevoPersonFromFile
Export-ModuleMember Get-EdlevoOrganization
Export-ModuleMember Get-EdlevoOrganizationStaff