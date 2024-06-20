function Get-Printers {
    $name = Read-Host "Enter the naming convention of your print servers"
    $num = Read-Host "Enter the number of print servers in your domain"
    $printers = @()

    for ($i = 1; $i -le $num; $i++) {
        $printers += (net view "\\$name$i" | Where-Object {$_ -match '\sPrint\s' }) -replace '\s\s+', ',' | ForEach-Object {[PSCustomObject]@{Server="$name$i"; Name=$_.Split(",")[0]; IP=$_.Split(",")[2]}}
    }

    return $printers
}

$printers = Get-Printers | Sort-Object Server, IP

$entries = @()
$progress = 0

foreach ($printer in $printers) {
    Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($printers.Count) printers." -PercentComplete ($progress/$printers.Count * 100)
    $progress++

    $printerModel = .\snmpwalk -r:$printer.IP -t:1 -os:1.3.6.1.2.1.25.3.2.1.3.0 -op:1.3.6.1.2.1.25.3.2.1.3.1 -q 2>$null
    $entries += [PSCustomObject]@{Server=$printer.Server; Name=$printer.Name; IP=$printer.IP; Model=$printerModel}
}
Write-Progress -Activity "Completed" -Completed

if ($entries) {$entries | Export-Csv -Path ".\Report.csv" -Encoding UTF8 -NoTypeInformation}
