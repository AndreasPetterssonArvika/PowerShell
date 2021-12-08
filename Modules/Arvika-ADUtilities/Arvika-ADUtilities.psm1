# Script copies attributes from one user to another
#
# Parameters:
# - $copyFromUser, the user to be copied from
# - $userToCopyTo, the user to be copied to, sAMAccountName
# - $userAttributes, the attributes to be copied
function Copy-ADUserAttributes {
    param (
        [Parameter(Mandatory,ValueFromPipeline)] $copyToUser,
        $copyFromUser,
        $userAttributes
    )

    # Låt stå tills skriptet bättre testat
    BREAK
    # Låt stå tills skriptet bättre testat

    $userToCopyFrom = '<sAMAccountName>'

    $userToCopyTo = '<sAMAccountName>'

    $attributes = @('description','title','department','company','manager','streetAddress','l','telephoneNumber','physicalDeliveryOfficeName')

    (Get-ADUser $userToCopyFrom -Properties $attributes).PSObject.Properties | foreach { $copiedAttribs=@{} } {$copiedAttribs.add($_.Name, $_.value) }

    $attribsToSet = @{}

    foreach ( $attribute in $attributes ) {
        if ( $copiedAttribs.$attribute) {
            $attribsToSet.Add($attribute,$copiedAttribs.$attribute)
        }
    }

    Set-ADUser $userToCopyTo -Replace $attribsToSet

}



