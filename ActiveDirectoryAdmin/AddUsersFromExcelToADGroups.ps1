# Requires the ImportExcel Module
# Adds users to groups based on an Excel Workbook with users on worksheets.
# Groups are named the same except the identifying number at the end.
# The corresponding worksheets are identified byt the same number.

$workbook = '<path to file>'
$worksheets = @('70','71','72','80','81','82','90','91','92')

ForEach ( $worksheet in $worksheets) {
    $groupName = "GRMinnebergsskolanPersonal$worksheet"
    $group = Get-ADGroup -Identity $groupName
    $userList = Import-Excel -Path $workbook -WorksheetName $worksheet -StartColumn 1 -EndColumn 1 -NoHeader
    $mailAddresses = $userList.P1
    foreach ($mail in $mailAddresses) { Get-ADUser -LDAPFilter "(mail=$mail)" | Add-ADPrincipalGroupMembership -MemberOf $group }
}