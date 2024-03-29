<#
Skriptet uppdaterar XS-grupperna baserat på Excelblad samlade i en mapp.
Testat och kört 2023-02-09

Uppgifter som ska finnas i AD för alla grupper
arvikaCOMSkolform
arvikaCOMEnhet
arvikaCOMKlass? Avdelning
arvikaCOMUpdateID. Ska vara "LS36330"
#>

[cmdletbinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$BaseFolder,          # Mappen med excelbladen
    [Parameter(Mandatory)][string]$AutomaticGroupOU,
    [Parameter()][ValidateSet('Groups','Members')][string[]]$UpdateType = ('Members'),       # Anger om grupperna eller medlemmarna ska uppdateras.
    [Parameter()][switch]$UpdateUserData
)

#$WhatIfPreference=$true

#Requires -modules ImportExcel
Import-Module ImportExcel -Verbose:$false

# Filter för uppdaterings ID
$arvikaCOMUpdateID = 'LS36330'
$UpdateIDFilter="(arvikaCOMUpdateID=$arvikaCOMUpdateID)"
$arvikaCOMSkolform='FSK'
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

    Write-Debug "Converting $TitleAbbr"

    $titleBSKAbbr='BSK'
    $titleBSK='Barnskötare'
    $titleFSKAbbr='FSK'
    $titleFSK='Förskollärare'
    $titleLF5Abbr='L F-5'
    $titleLF5='Lärare F-5'
    $titleSocPAbbr='Soc.p.'
    $titleSocP='Socialpedagog'
    $titleLF3Abbr='L F-3'
    $titleLF3='Lärare F-3'


    if ( $TitleAbbr -match $titleBSKAbbr ) {
        $retTitle = $titleBSK
    } elseif ( $TitleAbbr -match $titleFSKAbbr ) {
        $retTitle = $titleFSK
    } elseif ( $TitleAbbr -match $titleLF5Abbr ) {
        $retTitle = $titleLF5
    } elseif ( $TitleAbbr -match $titleSocPAbbr ) {
        $retTitle = $titleSocP
    } elseif ( $TitleAbbr -match $titleLF3Abbr ) {
        $retTitle = $titleLF3
    } elseif ( $TitleAbbr ) {
        Write-Verbose "Ohanterad förkortning $TitleAbbr, sätter Personal"
        $retTitle = 'Personal'
    } else {
        Write-Debug "Förkortning saknas, sätter Personal"
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

function New-XSGroup {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$GroupOU,
        [Parameter(Mandatory)][string]$Groupname,
        [Parameter(Mandatory)][string]$Groupmail,
        [Parameter(Mandatory)][string]$Klass,
        [Parameter(Mandatory)][string]$Enhet,
        [Parameter(Mandatory)][string]$Skolform,
        [Parameter(Mandatory)][string]$UpdateID
    )

    $groupInfo='<ignore/>'  # Data för gruppens info-attribut. Medför att den vanliga gruppuppdateringen inte körs.
    $XSGroupdescription='XS-grupp för förskolan, #36330'

    if ( $PSCmdlet.ShouldProcess("Skapar gruppen $Groupname med epost-adressen $GroupMail",$Groupname,"Skapar grupp") ) {
        New-ADGroup -Name $Groupname -DisplayName $Groupname -SamAccountName $Groupname -Description $XSGroupdescription -GroupCategory Security -GroupScope Global -Path $GroupOU -PassThru | Set-ADGroup -Replace @{mail="$Groupmail";info="$groupInfo";arvikaCOMUpdateID="$UpdateID";arvikaCOMKlass="$Klass";arvikaCOMEnhet="$Enhet";arvikaCOMSkolform="$Skolform"}
    }

}

$rektorPattern = '^Rektor [\w]{2,3}$'
$budgetPattern = '^Budget [\w]{2,3}$'
$identifierPattern = '^[\d]{8}-[\d]{4}$'

$drivePermissionString = 'XS'

# Slå upp alla Excelfiler i mappen

$worksheets = Get-ChildItem -Path $BaseFolder | Get-ExcelFileSummary | Select-Object -Property Excelfile,Path,Worksheetname

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
                    $candName = "$arvikaCOMSkolform$curAcr$curDept$groupXSIdentifier"
                    $candMail = "$arvikaCOMSkolform.$curAcr.$cleanedCurDept.$groupXSIdentifier@$arvikaDomain".ToLower()
                    Write-Verbose "Kandidatgrupp`: $candName $candMail"
                    if ( $curGroups.ContainsKey($candMail) ) {
                        # Gruppen är redan skapad, markera att den ska fortsätta finnas genom att byta ut $groupExists mot $keepGroup
                        Write-Verbose "Gruppen $candMail existerar redan bland kandidaterna"
                        $curGroups[$candMail] = $keepGroup
                    } else {
                        # Gruppen finns inte, skapar den
                        # Funktionen hanterar -Whatif internt
                        New-XSGroup -GroupOU $AutomaticGroupOU -Groupname $candName -Groupmail $candMail -Klass $curDept -Enhet $curWSName -Skolform $arvikaCOMSkolform -UpdateID $arvikaCOMUpdateID -WhatIf:$WhatIfPreference
                        # Lägg till den bland de befintliga grupperna och sätt att den ska behållas
                        $curGroups[$candMail]=$keepGroup
                        
                    }
                }
            }

        }
    }

    # Loopa igenom befintliga grupper och ta bort de som inte ska vara kvar
    # Detta är de grupper som fortfarande har $groupExists och inte fått det ändrat
    # till $keepGroup i dictionaryn
    Write-verbose "Går igenom listan över befintliga grupper för att ta bort de som inte hittats bland kandidaterna"
    foreach ( $mail in $curGroups.Keys ) {
        if ( $curGroups[$mail] -match $keepGroup ) {
            # Gruppen ska finnas kvar
            Write-Verbose "Behåller gruppen $mail"
        } else {
            # Gruppen ska tas bort
            Write-Verbose "Gruppen $mail finns på listan över grupper att ta bort"
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
            Write-Output "`nRektor`: $curWSName"
            Write-Verbose 'Rektorsbladet'
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
                
                
                # Bearbeta bara om det både finns ett ID och en uttycklig avdelning
                if ( ( $curID -match $identifierPattern ) -and ( $curDept -match '^[\w]{1,}' ) ) {
                    # Matchar identifierare och konverterar till ID12
                    $curClearTitle = Get-ClearTitleFromAbbr -TitleAbbr $curTitleAbbr                
                    $curID12 = ConvertTo-IDKey12 -IDKey13 $curID
                    $inputUsers[$curID12] = @{}
                    $inputUsers[$curID12]['Unit'] = $curWSName
                    $inputUsers[$curID12]['Dept'] = $curDept
                    $inputUsers[$curID12]['Title'] = $curClearTitle
                    if ( $curDriveInput -match $drivePermissionString ) {
                        Write-Debug "Har behörighet"
                        $inputUsers[$curID12]['XS'] = $true

                    } else {
                        Write-Debug "Har inte behörighet"
                        $inputUsers[$curID12]['XS'] = $false
                    }
                    
                    $testDept = $inputUsers[$curID12].Dept
                    $testTitle = $inputUsers[$curID12].Title
                    $testPermission = $inputUsers[$curID12].XS
                    Write-Debug "Hittade data för personal`: $curID12, $testDept, $testTitle, $testPermission"

                    # Hämta motsvarande användarnamn från Active Directory och lagra.
                    Write-Debug "Hämtar användarnamnet för $curID12"
                    $ldapfilter = "(personNummer=$curID12)"
                    $curUsername = Get-ADUser -LDAPFilter $ldapfilter | Select-Object -ExpandProperty sAMAccountName
                    $inputUsers[$curID12]['Username'] = $curUsername
                    # Uppdatera användardata baserat på vad som hittats
                    if ( $UpdateUserData ) {
                        if ( $PSCmdlet.ShouldProcess( "Uppdaterar data för $curUsername",$curUsername,"Uppdatera data")) {
                            Get-ADUser -LDAPFilter $ldapfilter | Set-ADUser -Replace @{title="$curClearTitle"}
                        }
                    }
                }
                
            }

            # Här ska gruppmedlemsskapen skötas
            Write-Debug 'Uppdatering av gruppmedlemsskap startar'

            # Slå upp befintliga XS-grupper
            # Skapa ett filter baserat på enhet och arvikaCOMUpdateID
            $curGroupFilter="(&$UpdateIDFilter(arvikaCOMEnhet=$curWSName))"
            $curGroups = Get-ADGroup -LDAPFilter $curGroupFilter -Properties arvikaCOMKlass,arvikaCOMEnhet,arvikaCOMSkolform
            foreach ( $group in $curGroups ) {
                $curDept = $group.arvikaCOMKlass
                Write-Debug "Hanterar gruppen för $curDept"

                # Hämta nuvarande användare ur gruppen till en dictionary
                $curUsers = @{}
                $group | Get-ADGroupMember | ForEach-Object { $curUsername = $_.sAMAccountName; $curUsers[$curUsername]=$curDept }
                if ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) {
                    foreach ( $key in $curUsers.Keys ) {
                        Write-Debug "Användaren $key finns i gruppen för $curDept"
                    }
                }

                # Hämta motsvarande användare ur användarunderlaget
                $curInputUsers = @{}
                $inputUsers.GetEnumerator().Where{ ($_.Value.Dept -eq $curDept) -and ( $_.Value.XS -eq $true ) } | ForEach-Object { $curUserName = $_.Value.Username; $curInputUsers[$curUserName]='inputuser' }

                if ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) {
                    foreach ( $key in $curInputUsers.Keys ) {
                        Write-Debug "Användaren $key ska finnas i gruppen för $curDept"
                    }
                }

                # Skapa hashtables med användare att lägga till resp ta bort
                $usersToAdd = Compare-HashtableKeys -Data $curInputUsers -Comp $curUsers
                $usersToRemove = Compare-HashtableKeys -Data $curUsers -Comp $curInputUsers

                # Lägg till och ta bort användare ur gruppen
                
                if ( $usersToAdd.Count -gt 0 ) {
                    if ( $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent ) {
                        Write-Verbose "Användare som läggs till"
                        $usersToAdd.Keys | Write-Verbose
                    }
                    if ( $PSCmdlet.ShouldProcess( "Lägger till användare i gruppen $group",$group,"Lägg till användare" ) ) {
                        $usersToAdd.Keys | Get-ADUser | Add-ADPrincipalGroupMembership -MemberOf $group
                    }
                }
                
                if ($usersToRemove.Count -gt 0 ) {
                    if ( $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent ) {
                        Write-Verbose "Användare som tas bort"
                        $usersToRemove.Keys | Write-Verbose
                    }
                    if ( $PSCmdlet.ShouldProcess( "Tar bort användare ur gruppen $group",$group,"Ta bort användare" ) ) {
                        $usersToRemove.Keys | Get-ADUser | Remove-ADPrincipalGroupMembership -MemberOf $group -Confirm:$false
                    }
                }
                
                
            }
        }
    }
    #>
    
}