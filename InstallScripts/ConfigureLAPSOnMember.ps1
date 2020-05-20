# This script is used for installing and configuring LAPS on a member server.
# Before this script runs the separate ConfigureLAPSMinimalDCScript.ps1 should be run on a Domain Controller

# Prevent running the entire script by mistake
BREAK

# Check if LAPS is installed on this computer
# See LAPS Detailed Technical Specification of Operational Procedures for file path info
$LAPSDLL =  "$env:ProgramFiles\LAPS\AdmPwd.Utils.dll"
if( -Not (test-path $LAPSDLL)) {
    #Not installed, quit script
    Write-Output 'LAPS doesnt seem to be installed on this computer'
    Write-Output "File $LAPSDLL is missing"
    BREAK
} else {
    Write-Output 'LAPS seems to be installed, starting configuration.'
}

# General variables
$DomainController = 'DC01'
$LAPSShareFolderLocation = 'C:\'
$LAPSShareFolderName = 'LAPS'
$LAPSShareFolderPath = "$LAPSShareFolderLocation$LAPSShareFolderName"
$LAPSShare = 'LAPS'

# The folder where the install files are located
$LAPSFilesPath = "$env:USERPROFILE\Documents"

######################
#                    #
# Import LAPS module #
#                    #
######################

# Import PS-module for LAPS
Import-Module AdmPwd.PS

###########################
#                         #
# Set up share with files #
#                         #
###########################

# Create folder to share
New-Item -Path "$LAPSShareFolderLocation\" -Name $LAPSShareFolderName -ItemType "Directory"

# If needed, set permissions for Security Principal on the folder
$acl = Get-Acl $LAPSShareFolderPath
$securityPrincipal = 'VIAMONSTRA\admminy'
$newACE = $securityPrincipal,'ReadAndExecute','ContainerInherit,ObjectInherit','None','Allow'
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $newACE
$acl.SetAccessRule($AccessRule)
$acl | Set-Acl $LAPSShareFolderPath

# Copy files to share
Copy-Item -Path "$LAPSFilesPath\LAPS.x*.msi" -Destination $LAPSShareFolderPath

# Share the folder
New-SmbShare -Name $LAPSShare -Path $LAPSShareFolderPath -FullAccess "Everyone"

#####################################
#                                   #
# LAPS User group and AD delegation #
#                                   #
#####################################

# Varibles for LAPS Group (users that will have access to LAPS)
$LAPSGroupName = 'Domain ViaMonstra LAPS'
$LAPSGroupDescription = 'LAPS Users'
$LAPSGroupPath = 'OU=Security Groups,OU=Internal IT,OU=ViaMonstra,DC=corp,DC=viamonstra,DC=com'

# Create LAPS group and add user
New-ADGroup -Description:$LAPSGroupDescription -GroupCategory:"Security" -GroupScope:"Global" -Name:$LAPSGroupName -Path:$LAPSGroupPath -SamAccountName:$LAPSGroupName -Server:$DomainController
Add-ADGroupMember -Identity $LAPSGroupName -Members 'admminy'

# OU with managed computers
$ManagedOU = 'OU=Workstations,OU=ViaMonstra,DC=corp,DC=viamonstra,DC=com'

# Check who has permission by default
$PermissionListOnInstall = 'LAPSPermissions_before_install.txt'
Find-AdmPwdExtendedRights -Identity VIAMONSTRA | Format-List | Out-File -FilePath "$LAPSFilesPath\$PermissionListOnInstall"

# Set Permissions for machines
Set-AdmPwdComputerSelfPermission -OrgUnit $ManagedOU

# Set permissions for LAPS group
Set-AdmPwdReadPasswordPermission -OrgUnit $ManagedOU -AllowedPrincipals $LAPSGroupName
Set-AdmPwdResetPasswordPermission -OrgUnit $ManagedOU -AllowedPrincipals $LAPSGroupName

# Check who has permission after setting out permissions
$PermissionListAfterUpdate = 'LAPSPermissions_after_install.txt'
Find-AdmPwdExtendedRights -Identity VIAMONSTRA | Format-List | Out-File -FilePath "$LAPSFilesPath\$PermissionListAfterUpdate"

# Script done