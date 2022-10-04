<#
Exemplet hämtar appar ur registret och plockar stegvis ut dem i tre olika variabler
#>

# Skapa en array och fyll den med värden ur registret
$apps = @()
$apps += Get-ItemProperty "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" # 32 Bit
$apps += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"             # 64 Bit

# Strängar för matchning med Where-Object
$publisher = 'Microsoft*'
$vcDisplayName = "[\w]*Visual C\+\+*"
$platformPattern = "[\w]*x64[\w]*"

# Hämta alla MS VC++ x64 men lägg alla MS-appar och MSVCpp-appar i variabler efter vägen som sen kan användas på annat håll
$MSVCpp_x64 = $apps | Where-Object { $_.Publisher -match $publisher } -OutVariable MSApps | Where-Object { $_.DisplayName -match $vcDisplayName } -OutVariable MSVCpp | Where-Object { $_.DisplayName -match $platformPattern } | Select-Object -Property DisplayName,DisplayVersion,InstallDate

# Hämta antalet från respektive variabel
$numMSApps = $MSApps | Measure-Object | Select-Object -ExpandProperty count
$numMSVCpp = $MSVCpp | Measure-Object | Select-Object -ExpandProperty count
$numMSVCpp_x64 = $MSVCpp_x64 | Measure-Object | Select-Object -ExpandProperty count

Write-Host "Antal Microsoft-appar: $numMSApps"
Write-Host "Antal Microsoft Visual C++: $numMSVCpp"
Write-Host "Antal Micrsoft Visual C++ x64: $numMSVCpp_x64"

# Lista detaljer om installerade Microsoft VIsual C++ x64
$MSVCpp_x64 | Select-Object -Property DisplayName,DisplayVersion,InstallDate