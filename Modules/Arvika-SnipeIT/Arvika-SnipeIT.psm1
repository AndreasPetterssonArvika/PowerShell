<#
Modulen hanterar diverse uppdateringar mellan Snipe-IT och andra system
#>

#Requires -Modules ImportExcel

<#
Funktionen hämtar uppdateringsdata från Active Directory
baserat på indata som användarnamn i form av epost-adresser.
#>
function New-SITUserUpdateFileFromAD {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory,ValueFromPipeline)]$Username,
        [string[]][Parameter(Mandatory)]$Attributes,
        [string][Parameter(Mandatory)]$UpdateFile
    )

    begin {
        # Skapa filen med rubriker
        $attrString = $Attributes -join ','
        $headerString = $Attributes -join '","'
        $headerString = "`"$headerString`",`"Username`""
        #Write-Verbose $Attributes
        Write-Verbose "Attribute string $attrString"
        Write-Verbose "Header string $headerString"
        $headerString | Out-File -Encoding utf8 -FilePath $UpdateFile

        
    }

    process {
        # Problemet är här

        # Lägg till rader med uppdateringsdata
        $outAttributes = "$attrString,$Username"
        Write-Verbose "Attributes to output $outAttributes"

        $userFilter = "(mail=$Username)"
        Write-Verbose "Current user to get $userFilter"
        Get-ADUser -LDAPFilter $userFilter -Properties $Attributes | Select-Object -Property $outAttributes | ConvertTo-Csv -Delimiter ',' -NoTypeInformation | Out-File -FilePath $UpdateFile -Encoding utf8 -Append
    }

    end {}

}

Export-ModuleMember New-SITUserUpdateFileFromAD