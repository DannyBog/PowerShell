$subnets = [ordered]@{
    "1.1.1.1/24" = "Site #1"
    "1.1.2.1/24" = "Site #2"
    "1.1.3.1/24" = "Site #3"
    "1.1.4.1/24" = "Site #4"
    "1.1.5.1/24" = "Site #5"
}

while ($choice = $subnets.GetEnumerator() | ForEach-Object {[PSCustomObject]@{Subnet=$_.Key; Site=$_.Value}} | Out-GridView -Title "Subnets" -OutputMode Single) {
    $subnet = $choice.Subnet.Split(".")[2]
    $result = 1..254 | ForEach-Object {Test-Connection 1.1.$subnet.$_ -Count 1 -AsJob} | Get-Job | Receive-Job -Wait | Select-Object @{n="IP"; e={$_.Address}}, @{n="Name"; e={if ($_.StatusCode -eq 0) {(Resolve-DnsName $_.Address).NameHost.Split(".")[0]}}}, @{n="Reachable"; e={if ($_.StatusCode -eq 0) {$true} else {$false}}} | Sort-Object -Property @{e="Reachable"; Descending=$true}, @{e="IP"; Descending=$false} | Out-GridView -Title "$($choice.Site)" -PassThru
    if ($result) {$result | Export-Csv -Path ".\$($choice.Site).csv" -Encoding UTF8 -NoTypeInformation}
}