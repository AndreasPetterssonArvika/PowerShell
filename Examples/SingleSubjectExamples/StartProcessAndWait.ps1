# Sätta skript på paus
# Starta processer
# Vänta på processer
#

# Pausa skriptet 5 sekunder
Start-Sleep -Seconds 5

# Starta processen notepad
# PassThru används för att få ut processobjektet som annars normalt sett inte returneras
$app = 'notepad.exe'
$myProcess = Start-Process -FilePath $app -PassThru

# Vänta på processen som startades
Wait-Process $myProcess.Id
Write-Host "Closed"