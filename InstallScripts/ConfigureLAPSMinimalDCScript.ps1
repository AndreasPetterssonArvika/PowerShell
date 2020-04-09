# This script is used for installing and configuring LAPS including the UI on a Domain Controller

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
$DomainController = $TargetServer

#########################################################
#                                                       #
# Import LAPS module and modify Active Directory Schema #
#                                                       #
#########################################################

# Import PS-module for LAPS
Import-Module AdmPwd.PS

# Modify AD Schema
Update-AdmPwdADSchema

# Policy settings not showing up in Group policy editor, should be under Computer/Policies/Adm Templates/LAPS
# Needs copying from local policy to domain policy

Copy-Item "$env:SystemRoot\PolicyDefinitions\AdmPwd.admx" "$env:SystemRoot\SYSVOL\domain\Policies\PolicyDefinitions"
Copy-Item "$env:SystemRoot\PolicyDefinitions\en-US\AdmPwd.adml" "$env:SystemRoot\SYSVOL\domain\Policies\PolicyDefinitions\en-US"