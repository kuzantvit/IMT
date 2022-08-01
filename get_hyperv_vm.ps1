
Get-SCVMMServer -ComputerName "inf-vmm-ha.infra.clouddc.ru" -UserRoleName "Read-Only DIB"
Get-SCVirtualMachine | ForEach-Object {
    $IPv4 = ($_ | Get-SCVirtualNetworkAdapter).ipv4Addresses
    $_ | Select-Object -Property VirtualMachineState, OperatingSystem, VirtualizationPlatform, CreationTime, Name, Tag, ComputerName, Description, @{N='ipv4Addresses';E={$IPv4}}
} | Export-Csv -Path "C:\audit_vm\report\vmm_report.csv" -Encoding UTF8


#Get-SCVirtualMachine -VMMServer "inf-vmm.test.ru" | select-Object -Property VirtualMachineState, OperatingSystem, VirtualizationPlatform, CreationTime, Name, ComputerNameString, ComputerName
# Get-SCVirtualMachine -VMMServer "inf-vmm.test.ru" -OnBehalfOfUser "domain\admin" -OnBehalfOfUserRole $a
#$b = Get-Content -Path C:\audit_vm\userrole.txt
