<#
Modulen hanterar diverse uppdateringar mellan Snipe-IT och andra system
#>

#Requires -Modules ImportExcel

<#
Funktionen hämtar uppdateringsdata från Active Directory
baserat på indata som användarnamn i form av epost-adresser.

Parametrar
$Username, användarens epost-adress. Anges manuellt eller via pipeline från t ex textfil
$Attributes, array med attribut som ska hämtas från Active Directory
$UpdateFile, namnet på utdatafilen som ska användas för import i Snipe-IT
$MissingUserFile, valfri parameter. Namnet på den fil där användare som inte hittas i AD ska skrivas.

#>
function New-SITUserUpdateFileFromAD {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory,ValueFromPipeline)]$Username,
        [string[]][Parameter(Mandatory)]$Attributes,
        [string][Parameter(Mandatory)]$UpdateFile,
        [string][Parameter()]$MissingUserFile
    )

    begin {
        # Skapa filen med rubriker
        $headerString = $Attributes -join '","'
        $headerString = "`"$headerString`",`"Username`""
        Write-Verbose "Header string $headerString"
        $headerString | Out-File -Encoding utf8 -FilePath $UpdateFile

        # Lägg till mail för att få en array att använda för uppslag mot Active Directory
        $lookupAttribs = $Attributes + 'mail'

        # Kontrollera om AD-attributet title ingår. Det behöver ändras till Job Title innan importen görs
        if ( $lookupAttribs -contains 'title' ) {
            Write-host "Attributen innehåller title. Kom ihåg att ändra det till `"Job Title`" i filen $UpdateFile innan importen görs i Snipe-IT"
        }
    }

    process {
        # Skriv den aktuella användaren till fil

        $ldapFilter = "(mail=$Username)"
        Write-Verbose "Current user to get $ldapFilter"
        
        $curUser = Get-ADUser -LDAPFilter $ldapFilter -Properties $lookupAttribs

        # Kontrollera om det finns ett resultat
        if ( $curUser -eq $null ) {
            # Användaren saknar mail
            Write-Host "Användaren $Username saknas"

            if ( $MissingUserFile ) {
                $Username | Out-File -FilePath $MissingUserFile -Encoding utf8 -Append
            }

        } else {

            $userRow = ''

            foreach ( $attrib in $lookupAttribs ) {
                Write-Verbose $curUser.$attrib
                $userRow += $curUser.$attrib + ','
            }
            
            $userRow = $userRow -replace ".$"
            Write-Verbose $userRow
            
            
            $userRow | Out-File -FilePath $UpdateFile -Encoding utf8 -Append

        }

    }

    end {}

}

Export-ModuleMember New-SITUserUpdateFileFromAD