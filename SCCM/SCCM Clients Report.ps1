$computername = Readh-Host "Enter the management point name"
$namespace = Read-Host "Enter the site code name"

$result = Get-WmiObject -ComputerName $computername -Namespace "Root\SMS\site_$namespace" -Class SMS_R_System | Select-Object Name, Client | Where-Object {($_.Client -eq 0) -or ($_.Client -eq $null)} | Sort-Object Name
if ($result) {$result | Export-Csv -Path ".\Report.csv" -Encoding UTF8 -NoTypeInformation}