# Script copies attributes from one user to another
#
# Parameters:
# - From-user, the user to be copied from
# - To-user, the user to be copied to
# - AttributesToCopy, the attributes to be copied

$attributes = @('description','title','department','company','manager','streetAddress','l','telephoneNumber','physicalDeliveryOfficeName')

(Get-ADUser pelle.testsson -Properties $attributes).PSObject.Properties | foreach { $hash=@{} } {$hash.add($_.Name, $_.value) }

$hash

$hash.Keys

foreach ( $key in $attributes ) {
    $tempval = $hash.$key
    "$key is $tempval"
}