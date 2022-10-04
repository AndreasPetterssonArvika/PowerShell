<#
Visar hur man jämför listor med varandra
#>

$listOne = @('Nisse','Kalle','Pelle')
$listTwo = @('Kalle','Pelle','Olle')

# Gemensamma namn för båda listorna
$listOne | Where-Object { $listTwo -contains $_ }

# Namn som bara finns i listOne
$listOne | Where-Object { $listTwo -notcontains $_ }