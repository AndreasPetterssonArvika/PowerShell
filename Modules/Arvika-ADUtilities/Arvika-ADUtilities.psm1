# Script copies attributes from one user to another
#
# Parameters:
# - $copyFromUser, the user to be copied from
# - $userToCopyTo, the user to be copied to, sAMAccountName
# - $userAttributes, the attributes to be copied
function Copy-ADAttributesFromUser {
    [cmdletbinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$copyFromUser,
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADAccount]$copyToUser,
        [Parameter()][string[]]$userAttributes = @('description','title','department','company','manager','streetAddress','l','telephoneNumber','physicalDeliveryOfficeName')
    )

    begin {
        Write-Verbose "Getting attributes from user $userToCopyFrom"
        (Get-ADUser -Identity $copyFromUser -Properties $userAttributes).PSObject.Properties | foreach { $copiedAttribs=@{} } {$copiedAttribs.add($_.Name, $_.value) }

        $attribsToSet = @{}

        foreach ( $attribute in $userAttributes ) {    # foreach triggers warning as an alias that should be changed. This instance of foreach is NOT an alias, rather MS has overloaded the meaning of foreach.
            if ( $copiedAttribs.$attribute) {
                $attribsToSet.Add($attribute,$copiedAttribs.$attribute)
            }
        }

        $jsonString = $attribsToSet | ConvertTo-Json
        Write-debug $jsonString
    }

    process {
            Write-Verbose "Setting attributes for users"
            Set-ADUser -Identity $copyToUser -Replace $attribsToSet -WhatIf:$WhatIfPreference
    }

    end {}

}

function Copy-ADGroupsFromUser {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$copyFromUser,
        [Parameter(Mandatory,ValueFromPipeline)][Microsoft.ActiveDirectory.Management.ADAccount]$copyToUser
    )

    begin {
        $groups = Get-ADUser -Identity $copyFromUser | Get-ADPrincipalGroupMembership | Out-GridView -PassThru
    }

    process {
        Get-ADUser -Identity $copyToUser | Add-ADPrincipalGroupMembership -MemberOf $groups
    }

    end {}
}

function Copy-ADGroupMembersToGroup {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)][string]$CopyFromGroup,
        [Parameter(Mandatory)][string]$CopyToGroup,
        [Parameter()][switch]$SelectFromGridView
    )

    if ( $SelectFromGridView) {
        Get-ADGroupMember -Identity $CopyFromGroup | Get-ADUser -Properties Name,mail,distinguishedName | Out-GridView -PassThru | Add-ADPrincipalGroupMembership -MemberOf $CopyToGroup
    } else {
        Get-ADGroupMember -Identity $CopyFromGroup | Add-ADPrincipalGroupMembership -MemberOf $CopyToGroup
    }

}

Export-ModuleMember Copy-ADAttributesFromUser
Export-ModuleMember Copy-ADGroupsFromUser
Export-ModuleMember Copy-ADGroupMembersToGroup