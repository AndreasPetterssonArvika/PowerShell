$cred = Get-Credential
$server = '<GCDS-server>'            # Server that runs GCDS, server.domain.com
$GCDSPath = '<Path to GCDS>'         # Path to GCDS on server, C:\Program Files\Google Cloud Directory Sync
$GCDSConfigFile = '<filename.xml>'   # Filename of GCDS-Config, filename.xml

$session = New-PSSession -ComputerName $server -Credential $cred

Invoke-Command -Session $session -ScriptBlock { cd $GCDSPath ; .\sync-cmd.exe -a -c .\$GCDSConfigFile }

Remove-PSSession $session