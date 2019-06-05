# Checks services on all domain controllers in domain

$dcs = (Get-ADDomain).ReplicaDirectoryServers
$svcs = "adws","dns","kdc","netlogon"
Get-Service -name $svcs -ComputerName $dcs | Sort Machinename | Format-Table -group @{Name="Computername";Expression={$_.Machinename.toUpper()}} -Property Name,Displayname,Status