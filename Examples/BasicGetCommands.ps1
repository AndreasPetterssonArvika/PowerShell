# Grundläggande Get Commands
# Kommandon som börjar med Get hämtar något som t ex Get-Process som hämtar alla aktiva processer och uppgifter om dem

# Filer i aktuell mapp
Get-ChildItem

# Filer i en annan mapp
Get-ChildItem -Path C:\temp

# Filer i mappen där skriptet körs
Get-ChildItem $PSScriptRoot

# Processer på den lokala datorn
Get-Process

# Slå upp alla processer för Chrome. Parametern -Name visar automatiskt fram tillgängliga processer
Get-process -Name chrome

# Slå upp alla services på datorn
Get-Service

# Slå upp alla services som börjar på sp och visa i GridView
Get-Service -Name sp* | Out-GridView

# Slå upp alla services som börjar på sp, visa i GridView. De som markeras skickas vidare och startas om. Hint: Print Spooler brukar gå bra att starta om...
Get-Service -Name sp* | Out-GridView -PassThru | Restart-Service

# Slå upp alla services och filtrera fram de som kör
Get-Service | Where-Object { $_.Status -eq 'Running' }