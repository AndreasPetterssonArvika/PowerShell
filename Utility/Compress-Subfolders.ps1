<#
Skriptet tar en mapp som parameter och komprimerar alla undermappar till enskilda zip-filer
Zip-filerna får samma namn som mappen som komprimeras och har filändelsen .zip
#>

[cmdletbinding()]
param(
    [string][Parameter(Mandatory)]$BaseFolder,
    [string][parameter(Mandatory)]$DestinationFolder
)
 
# Hämta alla undermappar
$subfolders = Get-ChildItem $BaseFolder -Directory
foreach ($s in $subfolders) {
 
$folderpath = $s.FullName
$foldername = $s.Name
$CompFileName = "$foldername.zip"
 
Compress-Archive -Path $folderpath -DestinationPath $DestinationFolder\$CompFileName
 
}