function Parse-KeyPackType {
    $type = $args[0]

    switch ($type) {
        "0" {return "Unknown"}
        "1" {return "Retail"}
        "2" {return "Volume"}
        "3" {return "Concurrent"}
        "4" {return "Temporary"}
        "5" {return "Open"}
        "6" {return "Not supported"}
    }
}

$name = Read-Host "Enter the name of your license server"
$licenses = Get-WmiObject -ComputerName $name -Class Win32_TSLicenseKeyPack | Where-Object {($_.KeyPackType -ne 0) -and ($_.KeyPackType -ne 4) -and ($_.KeyPackType -ne 6)}
$result = $licenses | Select-Object @{n="License Version and Type"; e={$_.ProductVersion + " - " + $_.TypeAndModel}}, @{n="License Program"; e={Parse-KeyPackType $_.KeyPackType}}, @{n="Total Licenses"; e={$_.TotalLicenses}}, @{n="Available"; e={$_.AvailableLicenses}}, @{n="Issued"; e={$_.IssuedLicenses}} | Out-GridView -Title "RD Licensing Manager" -PassThru
if ($result) {$result | Export-Csv -Path ".\CAL Usage.csv" -Encoding UTF8 -NoTypeInformation}