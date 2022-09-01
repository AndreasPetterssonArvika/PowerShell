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
        if ( $null -eq $curUser ) {
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

function Update-SITStudents {
    [cmdletbinding()]
    param (
        [string][Parameter(Mandatory)]$SnipeITUserExport,
        [string][Parameter(Mandatory)]$UserUpdateFileName,
        [string][Parameter(Mandatory)]$NewUserFileName,
        [string][Parameter(Mandatory)]$OldUserFileName,
        [string][Parameter(Mandatory)]$NameAndAcronymFile
    )

    # Skapa utdatafilerna
    "`"Username`",`"Department`",`"Location`"" | Out-File -FilePath $UserUpdateFileName -Encoding utf8
    "`"Username`",`"Department`",`"Location`",`"Company`",`"First name`",`"Last name`",`"EMail`",`"Job Title`"" | Out-File -FilePath $NewUserFileName -Encoding utf8

    # Hämta alla elever ur Snipe-IT-exporten
    $snipeitusers = Import-Csv -Delimiter ',' -Path $SnipeITUserExport | Where-Object { ( $_.Title -eq 'Elev' ) } | Select-Object -ExpandProperty Username

    # Lägg alla användarnamn från Snipe-IT i en dictionary
    $SITUsernames = @{}
    foreach ( $snipeituser in $snipeitusers ) {
        $SITUsernames.Add($snipeituser,'Elev')
    }

    $attributes = @('mail','arvikaCOMKlass','arvikaCOMEnhet')

    $ldapfilter = '(&(arvikaCOMSkolform=GR)(employeeType=student)(|(arvikaCOMKlass=4*)(arvikaCOMKlass=5*)(arvikaCOMKlass=6*)(arvikaCOMKlass=7*)(arvikaCOMKlass=8*)(arvikaCOMKlass=9*)))'
    Get-ADUser -LDAPFilter $ldapfilter -Properties $attributes | Get-SITNewAndUpdatedStudents -SITUserDict $SITUsernames -UserUpdateFileName $UserUpdateFileName -NewUserFileName $NewUserFileName -NameAndAcronymFile $NameAndAcronymFile

    Get-SITOldStudents -SITUserDict $SITUsernames -OldUserFile $OldUserFileName
    
}

function Get-SITNewAndUpdatedStudents {
    [cmdletbinding()]
    param (
        [Microsoft.ActiveDirectory.Management.ADObject][Parameter(Mandatory,ValueFromPipeline)]$ADUser,
        [hashtable][Parameter(Mandatory)]$SITUserDict,
        [string][Parameter(Mandatory)]$UserUpdateFileName,
        [string][Parameter(Mandatory)]$NewUserFileName,
        [string][Parameter(Mandatory)]$NameAndAcronymFile
    )

    begin {
        # Läs in dictionary för platser
        $unitNames = Get-Content -Path $NameAndAcronymFile | ConvertFrom-Csv -Delimiter ';'
        $AcronymTable = @{}
        foreach ( $row in $unitNames ) {
            $tName = $row.unitInputName
            $tAcr = $row.unitAcr    
            $AcronymTable.Add($tName,$tAcr)
        }
    }

    process {
        $curUsername = $ADUser.mail
        if ( $SITUserDict.ContainsKey($curUsername) ) {
            # Befintlig Snipe-IT-användare, lägg till i uppdateringsfilen för Snipe-IT
            $curDept = $ADUser.arvikaCOMKlass
            $curUnit = $ADUser.arvikaCOMEnhet

            # Skapa Location för eleven
            $tUnitAcr = $AcronymTable[$curUnit]
            if ( $curDept -match '^[4-6]{1}' ) {
                $ageBand = '4-6'
            } else {
                $ageBand = '7-9'
            }
            $curLocation = "GR $tUnitAcr $ageBand Elever"

            $updateRow = "$curUserName,$curDept,$curLocation"
            $updateRow | Out-File -FilePath $UserUpdateFileName -Encoding utf8 -Append

        } else {
            # Saknas i Snipe-IT, lägg till i filen med nya användare till Snipe-IT
            $curDept = $ADUser.arvikaCOMKlass
            $curUnit = $ADUser.arvikaCOMEnhet
            $givenName = $ADUser.givenName
            $SN = $ADUser.Surname


            # Skapa Location för eleven
            $tUnitAcr = $AcronymTable[$curUnit]
            if ( $curDept -match '^[4-6]{1}' ) {
                $ageBand = '4-6'
            } else {
                $ageBand = '7-9'
            }
            $curLocation = "GR $tUnitAcr $ageBand Elever"

            $updateRow = "$curUserName,$curDept,$curLocation,Arvika kommun,$givenName,$SN,$curUserName,Elev"
            $updateRow | Out-File -FilePath $NewUserFileName -Encoding utf8 -Append
        }
    }

    end {}

}

function Get-SITOldStudents {
    [cmdletbinding()]
    param (
        [hashtable][Parameter(Mandatory)]$SITUserDict,
        [string][Parameter(Mandatory)]$OldUserFile
    )

    "`"Username`",`"Department`",`"Location`",`"Job Title`"" | Out-File -FilePath $OldUserFile -Encoding utf8

    foreach ( $mail in $SITUserDict.Keys ) {
        # Kontrollera om användaren har ett konto i grundskolan
        $ldapfilter = "(&(mail=$mail)(arvikaCOMSkolform=GR))"
        $numUsers = Get-ADUser -LDAPFilter $ldapfilter | Measure-Object | Select-Object -ExpandProperty Count

        if ( $numUsers -lt 1 ) {
            # Användaren inte i grundskolan, lägg till i filen
            "$mail,Ingen klass,Fd elever,Fd elev i GR" | Out-File -FilePath $OldUserFile -Encoding utf8 -Append
        }
    }
}

Export-ModuleMember New-SITUserUpdateFileFromAD
Export-ModuleMember Update-SITStudents