# Import Active Directory module in a remote session and import the session to your local PowerShell

$server = '<server>'

$credential = Get-Credential
$session = New-PSSession -Name $server -Credential $credential

Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
Import-PSSession -Session $session -Module ActiveDirectory

# Don't forget to remove the session afterwards!
# Remove-PSSession -Session $session