<#
Skriptet visar hur Out-GridView fungerar genom att skicka alla lokala användare till Out-GridView
Namnen för de användare som markeras listas när man klickar OK nere till höger i listrutan.
#>

Get-LocalUser | Out-GridView -PassThru | Select-Object -ExpandProperty name