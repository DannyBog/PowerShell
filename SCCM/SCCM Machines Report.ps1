$computername = Readh-Host "Enter the management point name"
$namespace = Read-Host "Enter the site code name"

$query = @"
SELECT DISTINCT SMS_R_System.Name, SMS_G_System_COMPUTER_SYSTEM_PRODUCT.Vendor, SMS_G_System_COMPUTER_SYSTEM_PRODUCT.Version, SMS_G_System_COMPUTER_SYSTEM_PRODUCT.IdentifyingNumber, SMS_R_System.IPAddresses, SMS_R_System.MACAddresses, SMS_G_System_X86_PC_MEMORY.TotalPhysicalMemory, SMS_G_System_LOGICAL_DISK.Size, SMS_G_System_LOGICAL_DISK.FreeSpace, SMS_G_System_PROCESSOR.Name, SMS_G_System_OPERATING_SYSTEM.Caption
FROM SMS_R_System

INNER JOIN SMS_G_System_COMPUTER_SYSTEM_PRODUCT
ON SMS_G_System_COMPUTER_SYSTEM_PRODUCT.ResourceID = SMS_R_System.ResourceId

INNER JOIN SMS_G_System_X86_PC_MEMORY 
ON SMS_G_System_X86_PC_MEMORY.ResourceId = SMS_R_System.ResourceId

INNER JOIN SMS_G_System_LOGICAL_DISK
ON SMS_G_System_LOGICAL_DISK.ResourceID = SMS_R_System.ResourceID

INNER JOIN SMS_G_System_PROCESSOR
ON SMS_G_System_PROCESSOR.ResourceID = SMS_R_System.ResourceId

FULL JOIN SMS_G_System_TPM
ON SMS_G_System_TPM.ResourceID = SMS_R_System.ResourceId

INNER JOIN SMS_G_System_OPERATING_SYSTEM
ON SMS_G_System_OPERATING_SYSTEM.ResourceId = SMS_R_System.ResourceId

WHERE SMS_G_System_LOGICAL_DISK.DeviceID = "C:"
"@

$result = Get-WmiObject -ComputerName $computername -Namespace "Root\SMS\site_$namespace" -Query $query | Select-Object -Property @{n="Name"; e={$_.SMS_R_System.Name}}, @{n="Vendor"; e={$_.SMS_G_System_COMPUTER_SYSTEM_PRODUCT.Vendor}}, @{n="Version"; e={$_.SMS_G_System_COMPUTER_SYSTEM_PRODUCT.Version}}, @{n="Serial Number"; e={$_.SMS_G_System_COMPUTER_SYSTEM_PRODUCT.IdentifyingNumber}}, @{n="IP"; e={$_.SMS_R_System.IPAddresses}}, @{n="MAC Address"; e={$_.SMS_R_System.MACAddresses}}, @{n="RAM"; e={[int]($_.SMS_G_System_X86_PC_MEMORY.TotalPhysicalMemory * 1KB / 1GB)}}, @{n="Disk Size (GB)"; e={[math]::round($_.SMS_G_System_LOGICAL_DISK.Size * 1MB / 1GB, 1)}}, @{n="Disk Space (GB)"; e={[math]::round($_.SMS_G_System_LOGICAL_DISK.FreeSpace * 1MB / 1GB, 1)}}, @{n="Processor" ;e={$_.SMS_G_System_PROCESSOR.Name}}, @{n="Operating System" ;e={$_.SMS_G_System_OPERATING_SYSTEM.Caption}}
if ($result) {$result | Export-Csv -Path ".\Report.csv" -Encoding UTF8 -NoTypeInformation}