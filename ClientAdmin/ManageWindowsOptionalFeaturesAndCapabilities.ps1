# *********************
# *                   *
# * Optional Features *
# *                   *
# *********************

# Lista alla aktiverade Windows Optional features
Get-WindowsOptionalFeature -Online | Format-Table

# Lista alla aktiverade Windows Optional features
Get-WindowsOptionalFeature -Online | Where-Object { $_.State -eq 'Enabled' } | Format-Table

# ************************
# *                      *
# * Windows Capabilities *
# *                      *
# ************************

# Lista alla Windows Capabilities
Get-WindowsCapability -Online

# Lista alla RSAT-delar
Get-WindowsCapability -Online -Name rsat.* | Format-Table

# Installera RSAT DHCP
Get-WindowsCapability -Online -Name rsat.dhcp* | Add-WindowsCapability -Online

# Lista alla RSAT-delar som inte är installerade
Get-WindowsCapability -Online -Name rsat.* | Where-Object { $_.State -eq 'NotPresent' } | Format-Table

# Lista alla RSAT-delar som inte är installerade i en Grid-View och installera de som markeras
Get-WindowsCapability -Online -Name rsat.* | Where-Object { $_.State -eq 'NotPresent' } | Out-GridView -PassThru | Add-WindowsCapability -Online


# Vid problem med 20H2, testat
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
Restart-Service wuauserv

#Get-WindowsCapability -Online -Name rsat.dhcp* | Add-WindowsCapability -Online
#Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 1
Restart-Service wuauserv