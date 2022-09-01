<#

Funktionen skapar en importfil för Snipe-IT baserat på Ateas ShipmentAssetReport
Förutsätter tre fält i import-filen
- Serienr, ska innehålla serienummer
- Theftmark, ska innehålla vårt stöldmärkningsnummer från etiketten
- MAC-Adress Wifi, ska innehålla MAC-adressen i format med enbart hexadecimala siffror

#>

#Requires -Modules ImportExcel

# Dialogruta för att välja fil

Function Get-FileName {
    param (
        [string]$initialDirectory
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “Excel-filer (*.xlsx)| *.xlsx”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

# Dialogruta för att välja filnamn för att spara
Function Get-SaveFileName {  
    param (
        [string]$initialDirectory,
        [string]$DefaultFileName
    )
    [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “CSV-filer (*.csv)| *.csv”
    $OpenFileDialog.Title = "Välj fil"
    $OpenFileDialog.FileName = $DefaultFileName
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

function New-SnipeITLine {
    [cmdletbinding()]
    param (
        [parameter(ValueFromPipeline)]$asset,
        [Parameter(Mandatory)]
        [ValidateSet("Deployable","Ready to deploy")]
        [string]$AssetStatus,
        [string]$Location,
        [string]$outputFile
    )

    begin {}

    process {
        #$asset
        $tManufacturer = "Dell"
        $tModelName = "Chromebook 3100"
        $tModelNo="1WCWD"
        $tCategory = "Chromebook"
        $tSerial = $asset.Serienr
        $tAID = $asset.Theftmark
        $tMac = [string]$asset."MAC-Adress Wifi" -replace '..(?!$)','$&:' | ForEach-Object { $_.toLower() }
        $tString = "`"$tManufacturer`",`"$tModelName`",`"$tModelNo`",`"$tSerial`",`"$tAID`",`"$tMac`",`"$tCategory`",`"$AssetStatus`",`"$Location`""
        $tString | Out-File -FilePath $outputFile -Append -Encoding UTF8
    }

    end {}

}

if ($psISE) {
    # Objektet finns, skriptet körs från ISE.
    # Hämta sökvägen från $psISE
    $basePath = Split-Path -Path $psISE.CurrentFile.FullPath
} else {
    # Alla andra fall, använd $PSScriptRoot
    $basePath = $PSScriptRoot
}

$now = get-date -Format yyyyMMdd_HHmm

$assetFile = Get-FileName -initialDirectory $basePath
$assets = Import-Excel -Path $assetFile
$location = "GR Admin"
$defaultFileName = "myImport_$now.csv"
$outputFile = Get-SaveFileName -initialDirectory $basePath -DefaultFileName $defaultFileName
$headers = "`"Manufacturer`",`"Model Name`",`"Model Number`",`"Serial Number`",`"Asset Tag`",`"MAC Address`",`"Category`",`"Status`",`"Location`""
$headers | Out-File -FilePath $outputFile -Encoding UTF8


$assets | New-SnipeITLine -outputFile $outputFile -AssetStatus Deployable -Location $location