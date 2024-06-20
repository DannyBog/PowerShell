$dn = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName
$servers = Get-ADComputer -Filter * -SearchBase "OU=Servers,$dn"

$entries = @()
$progress = 0

foreach ($server in $servers) {
    Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($servers.Count) servers." -PercentComplete ($progress/$servers.Count * 100)
    $progress++

    $instances = (reg query "\\$($server.Name)\HKLM\SOFTWARE\Microsoft\Microsoft SQL Server" /v "InstalledInstances" 2>$null) -replace "\s\s+", ","
    if ($instances) {$instances = $instances.Split(",")[5]}

    foreach ($instance in $instances) {
        $name = (reg query "\\$($server.Name)\HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL" /se "#" 2>$null) -replace "\s\s+", ","
        $name = $name.Split(",")[5]

        $edition = (reg query "\\$($server.Name)\HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\$name\Setup" /v "Edition" 2>$null) -replace "\s\s+", ","
        $version = (reg query "\\$($server.Name)\HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\$name\Setup" /v "Version" 2>$null) -replace "\s\s+", ","

        $entries += [PSCustomObject]@{Server=$($server.Name); Edition=$edition.Split(",")[5]; Version=$version.Split(",")[5]};
    }
}
Write-Progress -Activity "Completed" -Completed

$result = $entries | Sort-Object Server | Out-GridView -Title "SQL Servers" -PassThru
if ($result) {$result | Export-Csv -Path ".\SQL Servers.csv" -Encoding UTF8 -NoTypeInformation}