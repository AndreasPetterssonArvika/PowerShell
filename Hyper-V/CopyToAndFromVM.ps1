<#
Kopiera filer till och från viruell maskin
Observera att det inte går att använda wildcards vid kopieringen
#>

# Kopiera en fil från host till VM
$VM = <Name of VM>
$SourceFile = <Local path to files on host>
$DestinationPath = <Local path to files on VM>

Copy-VMFile -FileSource Host -Name $VM -SourcePath $SourceFile -DestinationPath $DestinationPath

# Skriva över filer kräver växeln -Force

Copy-VMFile -FileSource Host -Name $VM -SourcePath $SourceFile -DestinationPath $DestinationPath -Force

<#
Kopiera flera filer
Eftersom kommandot Copy-VMFile inte accepterar wildcards kan man använda följande konstruktion som slår upp alla filer
i $SourceFiles och gör Copy-VMFile för var och en av dem
#>
$SourceFiles = <Local path to files on host>\*.*
Get-ChildItem -LiteralPath $SourceFiles | ForEach-Object { Copy-VMFile -FileSource Host -Name $VM -DestinationPath -SourcePath $_.FullName } 