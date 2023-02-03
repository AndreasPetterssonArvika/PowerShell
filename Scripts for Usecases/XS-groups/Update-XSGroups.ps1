<#
Skriptet uppdaterar XS-grupperna baserat på Excelblad samlade i en mapp.

Uppgifter som ska finnas i AD för alla grupper
arvikaCOMSkolform
arvikaCOMEnhet
arvikaCOMKlass? Avdelning
arvikaCOMUpdateID. Ska vara "LS36330"
#>

[cmdletbinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$BaseFolder,          # Mappen med excelbladen
    [Parameter()][ValidateSet('Groups','Members')][string[]]$UpdateType = ('Members')       # Anger om grupperna eller medlemmarna ska uppdateras.
)

$WhatIfPreference=$true

#Requires -modules ImportExcel
Import-Module ImportExcel

# Filter för uppdaterings ID
$UpdateIDFilter='(arvikaCOMUpdateID=LS36330)'
$arvikaCOMSKolform='FSK'
$arvikaDomain='arvika.com'
$groupXSIdentifier='XS'
$groupExists='exist'
$keepGroup='keep'

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

function ConvertTo-ANCAlfaNumeric {
    [cmdletbinding()]
    param(
        [string][Parameter(Mandatory,ValueFromPipeline)]$myString
    )

    # Byt ut icke alfanumeriska tecken
    $myString = $myString -replace '[^\p{L}\p{Nd}]', ''

    # Byt ut diverse diakritiska tecken
    # creplace är case sensitive
    $myString = $myString -creplace '[\u00C0-\u00C6]','A'
    $myString = $myString -creplace '[\u00E0-\u00E6]','a'
    $myString = $myString -creplace '[\u00C7]','C'
    $myString = $myString -creplace '[\u00E7]','c'
    $myString = $myString -creplace '[\u00C8-\u00CB]','E'
    $myString = $myString -creplace '[\u00E8-\u00EB]','e'
    $myString = $myString -creplace '[\u00CC-\u00CF]','E'
    $myString = $myString -creplace '[\u00EC-\u00EF]','e'
    $myString = $myString -creplace '[\u00D0]','D'
    $myString = $myString -creplace '[\u00F0]','d'
    $myString = $myString -creplace '[\u00D1]','N'
    $myString = $myString -creplace '[\u00F1]','n'
    $myString = $myString -creplace '[\u00D2-\u00D8]','O'
    $myString = $myString -creplace '[\u00F2-\u00F8]','o'
    $myString = $myString -creplace '[\u00D9-\u00DC]','U'
    $myString = $myString -creplace '[\u00F9-\u00FC]','u'
    $myString = $myString -creplace '[\u00DD]','Y'
    $myString = $myString -creplace '[\u00FD]','y'

    return $myString
}

function Get-ClearTitleFromAbbr {
    [cmdletbinding()]
    param(
        [string][parameter()]$TitleAbbr
    )

    Write-Verbose "Converting $TitleAbbr"

    $titleBSKAbbr='BSK'
    $titleBSK='Barnskötare'
    $titleFSKAbbr='FSK'
    $titleFSK='Förskollärare'
    $titleLF5Abbr='L F-5'
    $titleLF5='Lärare F-5'

    if ( $TitleAbbr -match $titleBSKAbbr ) {
        $retTitle = $titleBSK
    } elseif ( $TitleAbbr -match $titleFSKAbbr ) {
        $retTitle = $titleFSK
    } elseif ( $TitleAbbr -match $titleLF5Abbr ) {
        $retTitle = $titleLF5
    } elseif ( $TitleAbbr ) {
        Write-Host "Ohanterad förkortning $TitleAbbr"
        $retTitle = 'Personal'
    } else {
        Write-verbose "Förkortning saknas"
        $retTitle = 'Personal'
    }

    return $retTitle
}

# Dialogruta för att välja fil
Function Get-FileName {
    param (
        [string]$InitialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $InitialDirectory
    $OpenFileDialog.filter = “CSV-filer (*.csv)| *.csv”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Compare-HashtableKeys {
    <#
    Funktionen jämför hashtables.
    Utan switchen $CommonKeys kommer alla nycklar ur $data som saknas i $comp att returneras
    som nycklar i en ny hashtable, med värdet diff för alla nycklar
    Med switchen $CommonKeys kommer alla nycklar ur $data som också finns i $comp att returneras
    som nycklar i en ny hashtable, med värdet common för alla nycklar
    #>

    [cmdletbinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Data,
        [Parameter(Mandatory)][hashtable]$Comp,
        [Parameter()][switch]$CommonKeys
    )

    $return = @{}

    if ( $CommonKeys ) {
        foreach ( $key in $Data.Keys ) {
            if ( $Comp.ContainsKey( $key ) ) {
                # Gemensamma vÃ¤rden ska returneras, lÃ¤gg till i returen
                $return[$key]='common'
            } else {
                # GÃ¶r inget
            }
        }
    } else {
        foreach ( $key in $Data.Keys ) {
            if ( $Comp.ContainsKey( $key ) ) {
                # Diffen ska returneras, gÃ¶r inget
            } else {
                # Skilda vÃ¤rden ska returneras, lÃ¤gg till i returen
                $return[$key]='diff'
            }
        }
    }

    return $return
}

$rektorPattern = '^Rektor [\w]{2,3}$'
$budgetPattern = '^Budget [\w]{2,3}$'
$identifierPattern = '^[\d]{8}-[\d]{4}$'

$drivePermissionString = 'XS'

# Slå upp alla Excelfiler i mappen

$worksheets = Get-ChildItem -Path $BaseFolder | Get-ExcelFileSummary | Select-Object -Property Excelfile,Path,Worksheetname -First 5

# Om grupperna ska uppdateras
if ( $UpdateType -eq 'Groups' ) {
    # Hämta förkortningar till hashtable
    $unitAcrFile = Get-FileName -InitialDirectory .
    $unitAcrFileData = Import-Csv -Delimiter ';' -Path $unitAcrFile -Encoding utf8

    $unitAcrData = @{}

    foreach ( $unit in $unitAcrFileData ) {
        $unitAcrData[$unit.UnitDisplayName]=$unit.UnitAcr
    }

    # Lista namn och förkortningar för att se om de är korrekt importerade
    $unitAcrData
    $message = "Kontrollera att alla tecken i importen av förkortningar är korrekta"
    $message = $message + "`nTryck valfri tangent för att fortsätta eller Ctrl+C för att avbryta"
    Read-Host -Prompt $message

    # Hämta existerande grupper från Active Directory till hashtable
    # Ska ha mailadress som key och $groupExists som värde
    $message = 'Här ska existerande grupper hämtas och skrivas till dictionary'
    Read-Host -Prompt $message

    $curGroups = @{}
    Get-ADGroup -LDAPFilter $UpdateIDFilter -Properties mail | ForEach-Object { $curMail = $_.mail; $curGroups[$curMail]=$groupExists}

    # Loopa igenom alla hittade Worksheets
    foreach ( $sheet in $worksheets ) {
        $curWSName = $sheet.Worksheetname
        Write-Verbose "Aktuellt blad $curWSName"
        if ( $curWSName -match $rektorPattern ) {
            # Rektors blad, bearbeta inte
            Write-Verbose "Rektorsblad`: $curWSName"
        } elseif ( $curWSName -match $budgetPattern ) {
            # Budgetsammanställning, bearbeta inte
            Write-Verbose "Budgetblad`: $curWSName"
        } else {
            # Övriga blad
            Write-Verbose "Avdelning`: $curWSName"

            # Importera data och loopa igenom
            $curWBName = $sheet.Excelfile
            $curWBPath = $sheet.Path
            $curFile = "$curWBPath\$curWBName"
            $curWSContent = Import-Excel -Path $curFile -WorksheetName $curWSName -NoHeader -StartColumn 2 -EndColumn 6

            foreach ( $row in $curWSContent ) {
                # Hämta avdelning och identifierare
                [string]$curDept = $row.P1
                [string]$curID = $row.P2

                # Om raden innehåller personaldata och har ett riktigt avdelningsnamn
                if ( ( $curID -match $identifierPattern ) -and ( $curDept -match '[a-zA-Z]{1,}' ) ) {
                    [string]$curAcr = $unitAcrData[$curWSName]
                    Write-Verbose "Hittade data för grupper`: $curAcr, $curDept"
                    $cleanedCurDept = $curDept | ConvertTo-ANCAlfaNumeric
                    $candMail = "$arvikaCOMSKolform.$curAcr.$cleanedCurDept$groupXSIdentifier@$arvikaDomain".ToLower()
                    Write-Verbose "Kandidatgrupp`: $candMail"
                    if ( $curGroups.ContainsKey($candMail) ) {
                        # Gruppen är redan skapad, markera att den ska fortsätta finnas genom att byta ut $groupExists mot $keepGroup
                        Write-Verbose "Gruppen $candMail existerar redan bland kandidaterna"
                        $curGroups[$candMail] = $keepGroup
                    } else {
                        # Gruppen finns inte, skapar den
                        if ( $PSCmdlet.ShouldProcess("Skapar gruppen $candMail",$candMail,"Skapar grupp") ) {
                            Write-Verbose "Skapar gruppen $candMail"
                            Write-Host "Här ska det vara en funktion som skapar gruppen $candMail"
                            # Lägg till den bland de befintliga grupperna om den kunde skapas
                            # Sätt att den ska behållas
                            $curGroups[$candMail]=$keepGroup
                        }
                        
                    }
                }
            }

        }
    }

    # Loopa igenom befintliga grupper och ta bort de som inte ska vara kvar
    # Detta är de grupper som fortfarande har $groupExists och inte fått det ändrat
    # till $keepGroup i dictionaryn
    foreach ( $mail in $curGroups.Keys ) {
        if ( $curGroups[$mail] -match $keepGroup ) {
            # Gruppen ska finnas kvar
            Write-Verbose "Behåller gruppen $mail"
        } else {
            # Gruppen ska tas bort
            if ($PSCmdlet.ShouldProcess("Tar bort gruppen $mail",$mail,'Ta bort')) {
                Write-Verbose "Tar bort gruppen $mail"
                $ldapfilter = "(mail=$mail)"
                Get-ADGroup -LDAPFilter $ldapfilter | Remove-ADGroup -Confirm:$false
            }
        }
    }

} elseif ( $UpdateType -eq 'Members' ) {
    Write-Verbose "Uppdaterar medlemmar"

    #<#
    foreach ( $sheet in $worksheets ) {
        $curWSName = $sheet.Worksheetname
        if ( $curWSName -match $rektorPattern ) {
            Write-Verbose 'Rektorsbladet'
            Write-Host "Rektor`: $curWSName"
        } elseif ( $curWSName -match $budgetPattern ) {
            # Gör inget
            Write-Verbose 'Budgetbladet'
        } else {

            # Hämta data från Excelbladet
            Write-Verbose "Förskola`: $curWSName"
            $curWBName = $sheet.Excelfile
            $curWBPath = $sheet.Path
            $curFile = "$curWBPath\$curWBName"
            $curContent = Import-Excel -Path $curFile -WorksheetName $curWSName -NoHeader -StartColumn 2 -EndColumn 6
            
            $inputUsers = @{}

            foreach ( $row in $curContent ) {
                $curDept = $row.P1
                $curID = $row.P2
                $curDriveInput = $row.P4
                $curTitleAbbr = $row.P5
                
                
                if ( $curID -match $identifierPattern ) {
                    # Matchar identifierare och konverterar till ID12
                    $curClearTitle = Get-ClearTitleFromAbbr -TitleAbbr $curTitleAbbr                
                    $curID12 = ConvertTo-IDKey12 -IDKey13 $curID
                    $hasPermission = $false
                    $inputUsers[$curID12] = @{}
                    $inputUsers[$curID12]['Unit'] = $curWSName
                    $inputUsers[$curID12]['Dept'] = $curDept
                    $inputUsers[$curID12]['Title'] = $curClearTitle
                    if ( $curDriveInput -match $drivePermissionString ) {
                        Write-Verbose "Har behörighet"
                        $inputUsers[$curID12]['XS'] = $true

                    } else {
                        Write-Verbose "Har inte behörighet"
                        $inputUsers[$curID12]['XS'] = $false
                    }
                    #Write-Verbose "Hittade data för medlemmar`: $curWSName, $curDept, $curID12, $curClearTitle,$hasPermission"
                    $testDept = $inputUsers[$curID12].Dept
                    $testTitle = $inputUsers[$curID12].Title
                    $testPermission = $inputUsers[$curID12].XS
                    Write-Verbose "Hittade data för personal`: $curID12, $testDept, $testTitle, $testPermission"

                    # Hämta motsvarande användarnamn från Active Directory och lagra.
                    Write-Verbose "Hämtar användarnamnet för $curID12"
                    $ldapfilter = "(personNummer=$curID12)"
                    $curUsername = Get-ADUser -LDAPFilter $ldapfilter | Select-Object -ExpandProperty sAMAccountName
                    $inputUsers[$curID12]['Username'] = $curUsername
                    # Uppdatera användardata baserat på vad som hittats
                    Write-Host "Här ska det vara en funktion som uppdaterar användardata."
                    if ( $PSCmdlet.ShouldProcess( "Uppdaterar data för $curUsername",$curUsername,"Uppdatera data")) {
                        Get-ADUser -LDAPFilter $ldapfilter | Set-ADUser -Replace @{title="$curClearTitle"}
                    }
                }
                
            }

            # Här ska gruppmedlemsskapen skötas

            # Slå upp befintliga XS-grupper
            $curGroups = Get-ADGroup -LDAPFilter $UpdateIDFilter -Properties arvikaCOMKlass,arvikaCOMEnhet,arvikaCOMSkolform
            foreach ( $group in $groups ) {
                $curDept = $group.arvikaCOMKlass

                # Hämta nuvarande användare ur gruppen till en dictionary
                $curUsers = @{}
                $group | Get-ADGroupMember | ForEach-Object { $curUsername = $_.sAMAccountName; $curUsers[$curUsername]=$curDept }

                # Hämta motsvarande användare ur användarunderlaget
                $curInputUsers = @{}
                $inputUsers.GetEnumerator().Where{ $_.Value.Dept -eq $curDept } | ForEach-Object { $curUserName = $_.Value.Username; $curInputUsers[$curUserName]='inputuser' }

                # Skapa hashtables med användare att lägga till resp ta bort
                $usersToAdd = Compare-HashtableKeys -Data $curInputUsers -Comp $curUsers
                $usersToRemove = Compare-HashtableKeys -Data $curUsers -Comp $curInputUsers

                # Lägg till och ta bort användare ur gruppen
                if ( $PSCmdlet.ShouldProcess( "Lägger till användare i gruppen $group",$group,"Lägg till användare" ) ) {
                    $usersToAdd.Keys | Add-ADPrincipalGroupMembership -MemberOf $group
                }
                if ( $PSCmdlet.ShouldProcess( "Tar bort användare ur gruppen $group",$group,"Ta bort användare" ) ) {
                    $usersToRemove.Keys | Remove-ADPrincipalGroupMembership -MemberOf $group -Confirm:$false
                }
                
                
            }
        }
    }
    #>
    
}

