# Slå upp mailadress i AD- attributen mail och proxyAddresses
#
# Sökningen i proxyAddresses har wildcards för att tillåta
# formatering för olika tjänster

$ADUser = '<username>'
$ADDomain = '<maildomain>'
$ldapfilter = "(|(mail=$ADUser@$ADDomain)(proxyaddresses=*$ADUser@$ADDomain*))"
$ADattributes = @('mail','description')

Get-ADUser -LDAPFilter $ldapfilter -Properties $ADattributes

