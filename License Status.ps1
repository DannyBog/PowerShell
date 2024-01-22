$dn = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName
$servers = Get-ADComputer -Filter * -SearchBase "OU=Servers,$dn"

$entries = @()
$progress = 0

foreach ($server in $servers) {
    Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($servers.Count) servers." -PercentComplete ($progress/$servers.Count * 100)
    $progress++

    $wmi = Get-WmiObject -ComputerName $server.Name -ClassName SoftwareLicensingProduct -Filter "Name like 'Windows%'" 2>$null | Where-Object {$_.PartialProductKey} | Select-Object Description, LicenseStatus
    switch ($wmi.LicenseStatus) {
        0 {$wmi.LicenseStatus = "Unlicensed"; break}
        1 {$wmi.LicenseStatus = "Licensed"; break}
        2 {$wmi.LicenseStatus = "Out-Of-Box Grace Period"; break}
        3 {$wmi.LicenseStatus = "Out-Of-Tolerance Grace Period"; break}
        4 {$wmi.LicenseStatus = "Non-Genuine Grace Period"; break}
        5 {$wmi.LicenseStatus = "Notification"; break}
        6 {$wmi.LicenseStatus = "Extended Grace"; break}
    }

    $entries += [PSCustomObject]@{Server=$server.Name; Description=$wmi.Description; Status=$wmi.LicenseStatus}
}
Write-Progress -Activity "Completed" -Completed

$entries | Export-Csv -Path ".\Result.csv" -Encoding UTF8 -NoTypeInformation