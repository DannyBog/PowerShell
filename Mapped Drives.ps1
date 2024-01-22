$policy = Read-Host "Enter the name of the policy in charge of mapping drives"
[xml]$report = Get-GPOReport -Name $policy -ReportType XML

$entries = @()
$drives = $report.GPO.User.ExtensionData.Extension.DriveMapSettings.Drive
foreach ($drive in $drives) {
    $groups = $drive.Filters.FilterGroup
    $properties = $drive.Properties | Select-Object Letter, Label, Path
    $entries += [PSCustomObject]@{Letter=$properties.Letter; Label=$properties.Label; Groups=$groups.Name; Path=$properties.Path}
}

$result = $entries | Sort-Object Letter | Out-GridView -Title "Mapped Drives" -PassThru
if ($result) {$result | Export-Csv -Path ".\Mapped Drives.csv" -Encoding UTF8 -NoTypeInformation}