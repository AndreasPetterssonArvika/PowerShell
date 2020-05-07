# Copy files to and from VM

# Copy files from host to VM
$adminCredential = Get-Credential
$VMName = 'SKOLA01'
$session = New-PSSession -VMName $VMName -Credential $adminCredential
Copy-Item -ToSession $session -Path 'C:\Users\andreas.pettersson\Desktop\dbtemp\*' -Destination 'C:\Users\administrator\Documents' -Recurse
Remove-PSSession $session


# Copy files from VM to host
$adminCredential = Get-Credential
$session = New-PSSession -VMName $VMName -Credential $adminCredential
Copy-Item -FromSession $session -Path "C:\Users\administrator\Documents\*.ps1" -Destination 'C:\temp'
Remove-PSSession $session