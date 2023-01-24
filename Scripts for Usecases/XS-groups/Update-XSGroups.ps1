<#
Skriptet uppdaterar XS-grupperna baserat på Excelblad samlade i en mapp.
#>

[cmdletbinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$BaseFolder,          # Mappen med excelbladen
    [Parameter()][ValidateSet('Groups','Members')][string[]]$UpdateType = ('Members')       # Anger om grupperna eller medlemmarna ska uppdateras.
)

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
    # Ska ha mailadress som key och 'exist' som värde
    $message = 'Här ska existerande grupper hämtas och skrivas till dictionary'
    Read-Host -Prompt $message

    $curGroups = @{}
    #$groupExists

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
                        # Gruppen är redan skapad, markera att den ska fortsätta finnas
                        Write-Verbose "Gruppen $candMail existerar redan bland kandidaterna"
                        $curGroups[$candMail] = $keepGroup
                    } else {
                        # Gruppen finns inte, skapar den
                        Write-Verbose "Skapar gruppen $candMail"
                        Write-Host "Här ska det vara en funktion som skapar gruppen $candMail"
                        # Lägg till den bland de befintliga grupperna om den kunde skapas
                        $curGroups[$candMail]=$keepGroup
                    }
                }
            }

        }
    }

    # Loopa igenom befintliga grupper och ta bort de som inte ska vara kvar
    foreach ( $mail in $curGroups.Keys ) {
        if ( $curGroups[$mail] -match $keepGroup ) {
            # Gruppen ska finnas kvar
            Write-Verbose "Behåller gruppen $mail"
        } else {
            # Gruppen ska tas bort
            if ($PSCmdlet.ShouldProcess($mail,'Ta bort')) {
                Write-Verbose "Tar bort gruppen $mail"
                Write-Host "Här ska det vara en funktion som tar bort gruppen"
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
            Write-Verbose "Avdelning`: $curWSName"
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

                    # Hämta motsvarande mailadress från Active Directory
                    Write-Verbose "Hämtar mailadressen för $curID12"
                    #$inputUsers[$curID12]['mail'] = $curUserMail
                    # Uppdatera användardata baserat på vad som hittats
                    Write-Host "Här ska det vara en funktion som uppdaterar användardata."
                }
                
            }
                
        }
    }
    #>
    
}

