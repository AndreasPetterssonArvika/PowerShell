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
        $Attributes -join "," | Out-File -Encoding utf8 -FilePath $UpdateFile
    }

    process {
        # Lägg till rader med uppdateringsdata
        $userFilter = "(mail=$Username)"
        Get-ADUser -LDAPFilter $userFilter -Properties $Attributes | Select-Object -Property $Attributes,$Username | ConvertTo-Csv -Delimiter ',' | Out-File -FilePath $UpdateFile -Append
    }

    end {}

}

