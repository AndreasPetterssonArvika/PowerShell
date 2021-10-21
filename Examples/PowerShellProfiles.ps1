# Profiler i PowerShell

# Profiler är skript som körs när en PowerShell-host startar och hanterar inställningar specifikt för användare och host
# Vanliga hosts är PowerShell-konsolen, PowerShell ISE och Visual Studio Code med plugin för PowerShell

# De tillgängliga profilerna. Kör de här i de olika miljöerna där du arbetar med Powershell och se hur sökvägarna ändras.
$profile.AllUsersAllHosts
$profile.AllUsersCurrentHost
$profile.CurrentUserAllHosts
$profile.CurrentUserCurrentHost

# Visa innehåll i en profil
Get-Content $profile.AllUsersAllHosts               # Alla användare oavsett PowerShell-miljö på datorn
Get-Content $profile.AllUsersCurrentHost            # Inställningar för alla användare i t ex VS Code
Get-Content $profile.CurrentUserAllHosts            # Aktuell användare, oavsett host
Get-Content $profile.CurrentUserCurrentHost         # Specifikt för nuvarande användare och miljö

# Öppna profil-mapp
$profile.AllUsersCurrentHost | Split-Path | Invoke-Item

# Öppna ett av skripten (öppnas per default i notepad.exe)
$profile.CurrentUserAllHosts  | Invoke-Item

# Redigera skriptet för CurrentUserAllHosts och lägg till följande rad
# för att få ett hälsningsmeddelande när du startar upp en PowerShell-host på datorn
# Testa med PowerShell-konsolen och ISE och se att meddelandet dyker upp i båda
Write-Host "Welcome back`n"

# Fast vill man inte öppna filen kan man kan ju göra redigeringen från PowerShell också...  Och ta med den aktuella PowerShell-versionen...
$welcomeMessage = '"Welcome back!"'
"Write-Host $welcomeMessage`nGet-Host | Select-Object Version" | Out-File -FilePath $profile.CurrentUserAllHosts -Append -Encoding utf8
