#Windows Remote Server Admiistration Tools
Get-WindowsCapability -Name RSAT* -Online
Get-WindowsCapability -Name RSAT* -Online | Select-Object -Property Displayname,State
Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online