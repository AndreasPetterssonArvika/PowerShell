<#
Skriptet tar en mapp som parameter och komprimerar alla undermappar till enskilda zip-filer
#>

[cmdletbinding()]
param(
    [string][Parameter(Mandatory)]$BaseFolder
)
 
# Hämta alla undermappar
$subfolders = Get-ChildItem $BaseFolder -Directory
foreach ($s in $subfolders) {
 
$folderpath = $s.FullName
$foldername = $s.Name
 
Compress-Archive -Path $folderpath -DestinationPath $BaseFolder\$foldername
 
}