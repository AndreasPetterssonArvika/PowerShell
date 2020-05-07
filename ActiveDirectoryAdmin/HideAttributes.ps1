# Hiding attribute in Active Directory

# Builds on Active Directory Cookbook p 478

# See blogpost for info on flags
# http://www.frickelsoft.net/blog/?p=151
Enum searchFlags {
    IX = 1    # Set indexing for attribute
    PI = 2    # Set indexing per container, for one-level searches of a container with children
    AR = 4    # Set Ambiguous Name Resolution (ANR) for the attribute
    PR = 8    # Preserve attribute on deletion
    CP = 16   # Copy attribute when object is copied
    TP = 32   # Enable Touple indexing for attribute
    ST = 64   # Create subtree index
    CF = 128  # Confidential, will trigger an exception from object.setInfo() for some reason. Flag is set anyway
    NV = 256  # Enable auditing on attribute, setting bit disables auditing
    RO = 512  # Put attribute in filtered attribute set, used to exclude this attribute from being replicated to RODCs
}

[searchFlags]::CP

$strAttrName = 'Phone-Home-Primary'

$objString = "LDAP://cn=$strAttrName," + $root.schemaNamingContext

$root = [ADSI]'LDAP://RootDSE'
$objAttr = [ADSI]($objString)

# Unset all searchFlags, lab environment only
#$objAttr.put('searchFlags',0)
#$objAttr.setInfo()

#Get current value
$newValue = $objAttr.get('searchFlags')
$newValue

# Setting a bit in searchFlags
$newValue = $newValue -bor [searchFlags]::CF

# Clearing a bit in searchFlags
$newValue = $newValue -band -bnot [searchFlags]::CF
$newValue

# Put and set the new value
$objAttr.put('searchFlags',$newValue)
$objAttr.setInfo()

# Alternate, short version for setting the bit
$objAttr.put('searchFlags',$objAttr.get('searchFlags') -bor [searchFlags]::IX)
$objAttr.setInfo()

# Alternate, short version for clearing the bit
$objAttr.put('searchFlags',$objAttr.get('searchFlags') -band -bnot [searchFlags]::CP)
$objAttr.setInfo()

# Testing visibility, currently able to read the value. No restart of DC made after altering the Searchflags parameter
$userCred = Get-Credential
$adminCred = Get-Credential

$user = Get-ADUser -Credential $userCred -Identity 'kalle' -Properties 'homePhone'
$userCred.UserName + " " + $user.HomePhone

$user = Get-ADUser -Credential $adminCred -Identity 'kalle' -Properties 'homePhone'
$adminCred.UserName + " " + $user.HomePhone

$user.GivenName
$user.HomePhone