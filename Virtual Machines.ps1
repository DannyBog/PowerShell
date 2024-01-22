$vcsa = Read-Host "Enter the  VCenter Server Appliance server name (or IP)"
Connect-VIServer -Server "$vcsa" -Force

$entries = @()
$vms = Get-VM
foreach ($vm in $vms) {
    $vmAdditionalInfo = $vm | Get-View | Select-Object -ExpandProperty Guest

    $disk = $vmAdditionalInfo | Select-Object -ExpandProperty Disk
    $diskSpace = "{0:n2}" -f (($disk | Measure-Object -Property FreeSpace -Sum).Sum / 1GB)

    $snapshot = $vm | Get-Snapshot | Measure-Object | Select-Object Count
    $entries += [PSCustomObject]@{Name=$vm.Name; State=$vm.PowerState; OS=$vmAdditionalInfo.GuestFullName; Cores=$vm.NumCpu; "RAM (GB)"=$vm.MemoryGB; "Disk Size (GB)"=$vm.ProvisionedSpaceGB; "Disk Space (GB)"=$diskspace; "VMTools Version"=$vmAdditionalInfo.ToolsVersion; "VMTools Status"=$vmAdditionalInfo.ToolsStatus; Snapshots=$snapshot.Count}
}

$result = $entries | Sort-Object Name | Out-GridView -Title "Virtual Machines" -PassThru
if ($result) {$result | Export-Csv -Path ".\Virtual Machines.csv" -Encoding UTF8 -NoTypeInformation}