<#
Skriptet uppdaterar XS-grupperna baserat på Excelblad samlade i en mapp.
#>

[cmdletbinding()]
param(
    [string][Parameter(Mandatory)]$BaseFolder
)

#Requires -modules ImportExcel
Import-Module ImportExcel

$rektorPattern = '^Rektor [\w]{2,3}$'
$budgetPattern = '^Budget [\w]{2,3}$'
$identifierPattern = '^[\d]{8}-[\d]{4}$'

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
            $curTitleAbbr = $row.P5
            #<#
            if ( $curID -match $identifierPattern ) {
                # Matchar identifierare
                Write-Verbose "Hittade data`: $curDept, $curID, $curTitleAbbr"
            }
            #>
        }
        #>
    }
}
#>

# För varje fil, slå upp alla blad

# För varje blad, hämta alla rader