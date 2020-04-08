#Before running this script the appropriate version of LAPS should be installed on the server

# Prevent running the entire script by mistake
BREAK

# Check if LAPS is installed on this computer
# See LAPS Detailed Technical Specification of Operational Procedures for file path info
$LAPSDLL =  "$env:ProgramFiles\LAPS\AdmPwd.Utils.dll"
if( -Not (test-path $LAPSDLL)) {
    #Not installed, quit script
    echo 'LAPS doesnt seem to be installed on this computer'
    echo "File $LAPSDLL is missing"
    BREAK
} else {
    echo 'LAPS seems to be installed, starting configuration.'
}

# General variables
$TargetServer = "$env:COMPUTERNAME"
$DomainController = 'DC01'
$LAPSShareFolderLocation = 'C:\'
$LAPSShareFolderName = 'LAPS'
$LAPSShareFolderPath = "$LAPSShareFolderLocation$LAPSShareFolderName"
$LAPSShare = 'LAPS'

# The folder where the install files are located
$LAPSFilesPath = "$env:USERPROFILE\Documents"

#########################################################
#                                                       #
# Import LAPS module and modify Active Directory Schema #
#                                                       #
#########################################################

# Import PS-module for LAPS
Import-Module AdmPwd.PS

# Modify AD Schema
Update-AdmPwdADSchema

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

# Create batch files for install on client computers
# Note: These does not work currently. Install manually with default settings from share
# or use SCCM or similar

#$cmdline_x64 = "`"msiexec /i \\$targetserver\$LAPSShare\LAPS.x64.msi /quiet`""
#$cmdline_x86 = "`"msiexec /i \\$targetserver\$LAPSShare\LAPS.x86.msi /quiet`""
#$LAPSClientInstallx64 = "$LAPSShareFolderPath\LAPS_x64.bat"
#$LAPSClientInstallx86 = "$LAPSShareFolderPath\LAPS_x86.bat"

#$cmdline_x64 | Out-File -FilePath $LAPSClientInstallx64
#$cmdline_x86 | Out-File -FilePath $LAPSClientInstallx86

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
BREAK

###################
#                 #
# Troubleshooting #
#                 #
###################

# Policy settings not showing up in Group policy editor, should be under Computer/Policies/Adm Templates/LAPS
# Needs copying from local policy to domain policy

Copy-Item "$env:SystemRoot\PolicyDefinitions\AdmPwd.admx" "$env:SystemRoot\SYSVOL\domain\Policies\PolicyDefinitions"
Copy-Item "$env:SystemRoot\PolicyDefinitions\en-US\AdmPwd.adml" "$env:SystemRoot\SYSVOL\domain\Policies\PolicyDefinitions\en-US"

# Get password for computer
Get-AdmPwdPassword -ComputerName PC001