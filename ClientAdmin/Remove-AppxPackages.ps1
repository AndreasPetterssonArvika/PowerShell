# Removes a number of provisioned packages that installs when new users log on to the computer
# Currently not parameterized. Awaits the possibility of signing the code

# This file should contain one or more strings describing packages, one per row.
# Package names can have wildcards
$packageFile = '<path to file containing packages>'

ForEach ($package in Get-Content $packageFile)
{
    # Remove the provisioned package from the computer
    Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like "$package" } | Remove-AppProvisionedPackage -Online
}