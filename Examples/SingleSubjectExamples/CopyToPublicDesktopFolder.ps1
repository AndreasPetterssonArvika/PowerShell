$sourcePath = 'C:\Temp'
New-Item -ItemType Directory -Force -Path $sourcePath | Out-Null
$now = Get-Date -Format yyMMdd_HHmm
#$today
$copyFile = $sourcePath+ "\copyFile_$now.txt"
#$copyFile
$message = "Meddelande skapat $now"
$message | Out-File -FilePath $copyFile

$targetPath = $env:PUBLIC + '\Desktop'
#$targetPath

Copy-Item $copyFile -Destination $targetPath