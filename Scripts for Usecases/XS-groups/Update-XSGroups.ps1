<#
Skriptet uppdaterar XS-grupperna baserat på Excelblad samlade i en mapp.
#>

[cmdletbinding()]
param(
    [string][Parameter(Mandatory)]$BaseFolder
)

#Requires -modules ImportExcel
Import-Module ImportExcel


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

$rektorPattern = '^Rektor [\w]{2,3}$'
$budgetPattern = '^Budget [\w]{2,3}$'
$identifierPattern = '^[\d]{8}-[\d]{4}$'

$drivePermissionString = 'J'

# Slå upp alla Excelfiler i mappen

#Get-ChildItem -Path $BaseFolder | Get-ExcelFileSummary
$worksheets = Get-ChildItem -Path $BaseFolder | Get-ExcelFileSummary | Select-Object -Property Excelfile,Path,Worksheetname -First 5

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
        $curWBName = $sheet.Excelfile
        $curWBPath = $sheet.Path
        $curFile = "$curWBPath\$curWBName"
        Write-Verbose "Enhetsblad`:$curWSName"
        #$curWBPath
        $curContent = Import-Excel -Path $curFile -WorksheetName $curWSName -NoHeader -StartColumn 2 -EndColumn 6
        
        #<#
        foreach ( $row in $curContent ) {
            $curDept = $row.P1
            $curID = $row.P2
            $curDriveInput = $row.P4
            $curTitleAbbr = $row.P5
            
            #<#
            if ( $curID -match $identifierPattern ) {
                # Matchar identifierare
                $curClearTitle = Get-ClearTitleFromAbbr -TitleAbbr $curTitleAbbr                
                $curID12 = ConvertTo-IDKey12 -IDKey13 $curID
                $hasPermission = $false
                if ( $curDriveInput -match $drivePermissionString ) {
                    $hasPermission = $true
                }
                Write-Verbose "Hittade data`: $curDept, $curID12, $curClearTitle,$hasPermission"
            }
            #>
        }
        #>
    }
}
#>

# För varje fil, slå upp alla blad

# För varje blad, hämta alla rader