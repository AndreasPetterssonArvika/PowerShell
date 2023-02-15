<#
Enumerators för hashtables
#>

$userDict = @{}

$userID = '1'
$userName = 'nisse.hult'
$userDept = 'Fikabrigaden'
$userLoc = 'Arvika'
$userXS = $true

Add-UserToDict -UserDict $userDict -UserID $userID -UserName $userName -UserDept $userDept -UserLoc $userLoc -UserXS $userXS

$userID = '2'
$userName = 'kalle.karlsson'
$userDept = 'Fikabrigaden'
$userLoc = 'Eda'
$userXS = $false

Add-UserToDict -UserDict $userDict -UserID $userID -UserName $userName -UserDept $userDept -UserLoc $userLoc -UserXS $userXS

$userID = '3'
$userName = 'pelle.persson'
$userDept = 'Arbetare'
$userLoc = 'Arvika'
$userXS = $true

Add-UserToDict -UserDict $userDict -UserID $userID -UserName $userName -UserDept $userDept -UserLoc $userLoc -UserXS $userXS

$userID = '4'
$userName = 'olle.olsson'
$userDept = 'Arbetare'
$userLoc = 'Eda'
$userXS = $true

Add-UserToDict -UserDict $userDict -UserID $userID -UserName $userName -UserDept $userDept -UserLoc $userLoc -UserXS $userXS


$userDict.GetEnumerator().Where{ $_.Value.XS -eq $true }
$userDict.GetEnumerator().Where{ $_.Value.Dept -eq 'Fikabrigaden' }
$userDict.GetEnumerator().Where{ $_.Value.Dept -eq 'Fikabrigaden' } | ForEach-Object { $_.Value.Dept }
$userDict.GetEnumerator().Where{ $_.key -gt 2 }
$userDict.GetEnumerator().Where{ $_.key -gt 2 } | ForEach-Object { $_.Value.Dept }
$userDict.GetEnumerator().Where{ ($_.Value.Dept -eq 'Fikabrigaden') -and ($_.Value.Loc -eq 'Arvika') }
$userDict.GetEnumerator().Where{ ($_.Value.Dept -eq 'Fikabrigaden') -and ($_.Value.Loc -eq 'Arvika') }  | ForEach-Object { $_.Value.Name }

$userDict.Count