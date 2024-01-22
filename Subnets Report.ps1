$subnets = [ordered]@{
    "1.1.1.1/24" = "Site #1"
    "1.1.2.1/24" = "Site #2"
    "1.1.3.1/24" = "Site #3"
    "1.1.4.1/24" = "Site #4"
    "1.1.5.1/24" = "Site #5"
}

$entries = @()
$entries += $subnets.GetEnumerator() | ForEach-Object {[PSCustomObject]@{Subnet=$_.Key; Site=$_.Value}}

$result = @()
$progress = 0

foreach ($entry in $entries) {
    $progress++
    Write-Progress -Activity "Pinging Subnets ($($entry.Site))..." -Status "$progress out of $($entries.Count)." -PercentComplete ($progress/$entries.Count * 100)

    $subnet = $entry.Subnet.Split(".")[2]
    $result += 1..254 | ForEach-Object {Test-Connection 1.1.$subnet.$_ -Count 1 -AsJob} | Get-Job | Receive-Job -Wait | Select-Object @{n="Site"; e={$entry.Site}}, @{n="IP"; e={$_.Address}}, @{n="Name"; e={(Resolve-DnsName -Name $_.Address).NameHost.Split(".")[0]}}, @{n="Reachable"; e={if ($_.StatusCode -eq 0) {$true} else {$false}}}
}
Write-Progress -Activity "Completed" -Completed

if ($result) {$result | Sort-Object -Property @{e="Reachable"; Descending=$true}, @{e="IP"; Descending=$false} | Export-Csv -Path ".\Report.csv" -Encoding UTF8 -NoTypeInformation}