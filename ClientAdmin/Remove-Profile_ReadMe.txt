You can use it to remove all the profiles in the local computer:

Remove-Profile

To remove all the profiles, except a specific profile:
Remove-Profile -Exclude Administrator

To remove profiles last used in the past 90 days that would be deleted:
Remove-Profile -DaysOld 90 | Where-Object { $_.WouldBeRemoved }

To remove against a remote computer:
Remove-Profile -ComputerName myRemoteServer

To remove against a collection of remote computers, authenticating different credentials:
Remove-Profile -ComputerName $Computers -Credential $cred -DaysOld 90

To ignore the LastUseTime value in case the account never logged-on locally (e.g. an IIS ApplicationPool), use the IgnoreLastUseTime switch:
Remove-Profile -ComputerName WebServer01 -DaysOld 30 -IgnoreLastUseTime 

To really remove the profiles, use the -Remove switch:
Remove-Profile -DaysOld 90 -Remove