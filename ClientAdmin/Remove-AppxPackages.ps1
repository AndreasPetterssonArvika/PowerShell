# A skeleton script that removes provisioned app packages that installs when a new user logs on to the computer

ForEach ($package in Get-Content '<path to package list>')
{
    # Ta bort paketet från datorn
    Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like "$package*" } | Remove-AppProvisionedPackage -Online
}