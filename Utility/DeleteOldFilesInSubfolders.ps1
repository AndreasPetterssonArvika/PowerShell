$path = "<path to root folder>"
$daysOld = -30
$timeLimit = (Get-Date).AddDays($daysOld)

cd $path
$files = Get-ChildItem -Recurse -Directory | ForEach-Object { $_.GetFiles()} | Where-Object { $_.Extension -eq ".txt" -and $_.CreationTime -lt $timeLimit}
$files

foreach ($file in $files) {
    Write-host "Deleting file $file";
    #Remove-Item $file.FullName -WhatIf
    Remove-Item $file.FullName
}