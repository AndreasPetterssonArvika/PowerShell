# Script copies attributes from one user to another
#
# Parameters:
# - $userToCopyFrom, the user to be copied from, sAMAccountName
# - $userToCopyTo, the user to be copied to, sAMAccountName
# - $attributes, the attributes to be copied

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