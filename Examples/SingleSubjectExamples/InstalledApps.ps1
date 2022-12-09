<#
Enkelt exempel som visar hur man slår upp installerade appar på ett par olika sätt
Notera att jag använt Get-CimInstance istället för Get-WmiObject som är föråldrat.
Exempel med Get-VmiObject kan oftast enkelt översättas till Get-CimInstance

Det finns problem med den här metoden, se här:
https://xkln.net/blog/please-stop-using-win32product-to-find-installed-software-alternatives-inside/
#>

$name='Cortex*'
$version='7.8'
$vendor='Palo*'

# Alla appar som hittas, lägg märke till att det mycket väl kan saknas appar
Get-CimInstance -ClassName Win32_Product

# Antal appar som hittas
Get-CimInstance -ClassName Win32_Product | Measure-Object | Select-Object -ExpandProperty count

# Lista alla appar som matchar namnet
Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -match $name } |  Select-Object -Property name,version

# Lista alla appar från en specifik Vendor
Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Vendor -match $vendor } |  Select-Object -Property name,version,vendor

# Antal appar som matchar namn och har en version högre än eller lika med den angivna
Get-CimInstance -ClassName Win32_Product | Where-Object { ( $_.Name -match $name ) -and ( $_.Version -ge $version ) } | Measure-Object | Select-Object -ExpandProperty count

<#
Slå upp installerade appar ur registret.
Visar fler appar och arbetar avsevärt snabbare.

Det går även att hitta appar som är installerade per användare, men jag har inte lagt till några exempel än
#>

# Skapa en array och fyll den med värden ur registret
$apps = @()
$apps += Get-ItemProperty "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" # 32 Bit
$apps += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"             # 64 Bit

# 
$apps[3]

# Antal appar som hittades i registret
$apps.Count

# Installerade Microsoft-appar
$publisher = 'Microsoft*'
$apps | Where-Object { $_.Publisher -match $publisher } | Select-Object -Property DisplayName,DisplayVersion,InstallDate

# Finns Palo Alto Cortex med minst version 7.8.1
$name='Cortex*'
$version='7.8.1'
$publisher='Palo*'

$apps | Where-Object { ( $_.Publisher -match $publisher ) -and ( $_.DisplayVersion -ge $version ) } | Select-Object -Property DisplayName,DisplayVersion,InstallDate