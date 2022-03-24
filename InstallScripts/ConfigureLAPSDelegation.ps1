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

# The folder where the install files are located
$LAPSFilesPath = "$env:USERPROFILE\Documents"

######################
#                    #
# Import LAPS module #
#                    #
######################

# Import PS-module for LAPS
Import-Module AdmPwd.PS

#####################################
#                                   #
# LAPS User group and AD delegation #
#                                   #
#####################################

# Group for LAPS-administrators

$LAPSGroupName = 'LAPS administrators'

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