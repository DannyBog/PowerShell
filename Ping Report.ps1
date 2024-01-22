function Parse-Date {
    $date = $args[0]

    $day = $date.Split(" ")[0].Split("/")[0]
    $month = $date.Split(" ")[0].Split("/")[1]
    
    if ($month -gt 12) {  
        $day = $date.Split(" ")[0].Split("/")[1]
        $month = $date.Split(" ")[0].Split("/")[0]    
    }

    if ([int]$day -le 10 -and $day.Length -eq 1) {$day = [String]0 + $day}
    if ([int]$month -le 10 -and $month.Length -eq 1) {$month = [String]0 + $month}

    $year = $date.Split(" ")[0].Split("/")[2]

    $hour = $date.Split(" ")[1].Split(":")[0]
    if ([int]$hour -le 10 -and $hour.Length -eq 1) {$hour = [String]0 + $hour}
    $minute = $date.Split(" ")[1].Split(":")[1]
    if ([int]$minute -le 10 -and $minute.Length -eq 1) {$minute = [String]0 + $minute}
    $second = $date.Split(" ")[1].Split(":")[2]
    if ([int]$second -le 10 -and $second.Length -eq 1) {$second = [String]0 + $second}

    if ($date -match "PM") {
        $hour = [string]([int]$date.Split(" ")[1].Split(":")[0] + 12)
        $minute = $date.Split(" ")[1].Split(":")[1]
        if ([int]$minute -le 10 -and $minute.Length -eq 1) {$minute = [String]0 + $minute}
        $second = $date.Split(" ")[1].Split(":")[2]
        if ([int]$second -le 10 -and $second.Length -eq 1) {$second = [String]0 + $second}
    }

    $formattedDate = $day + "/" + $month + "/" + $year + " " + $hour + ":" + $minute + ":" + $second
    return [DateTime]::ParseExact($formattedDate, "dd/MM/yyyy HH:mm:ss", $null)
}

$computers = Get-Content -Path ".\computers.txt" | ForEach-Object {Test-Connection $_ -Count 1 -AsJob} | Get-Job | Receive-Job -Wait | Select-Object @{n="Name"; e={$_.Address}}, @{n="Reachable"; e={if ($_.StatusCode -eq 0) {"True"} else {"False"}}}

$entries = @()

foreach ($computer in $computers) {
    try {
        $computerObj = ""
        $computerObj = Get-ADComputer -Identity $computer.Name -Properties Description
    } catch {
        $name = $computer.Name
        $samName = "NOT IN DATABASE"
        $username = "NOT IN DATABASE"
        $lastChecked = "NOT IN DATABASE"
        $ip = "NOT IN DATABASE"
        $serialNum = "NOT IN DATABASE"
        $reachable = ""
    }

    if ($computerObj) {
        $description = $computerObj.Description
        $name = $computerObj.Name

        if ($description) {
            $samName = $computerObj.Description.Split("()")[1]
            $username =  $samName + " " + "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
            $lastChecked = ForEach-Object {Parse-Date $computerObj.Description.Split("()")[3]} 2>$null
            $ip = $computerObj.Description.Split("()")[5]
            $serialNum = if ($computer.Reachable -eq "True") {Get-WmiObject -ComputerName $computerObj.Name -Class Win32_Bios 2>$null | Select-Object -ExpandProperty SerialNumber}
            $reachable = $computer.Reachable
        } else {
            $samName = ""
            $username = ""
            $lastChecked = ""
            $ip = ""
            $serialNum = ""
            $reachable = $computer.Reachable
        }   
    }

    $entries += [PSCustomObject]@{Name=$name; Username=$username; "Last Checked"=$lastChecked; IP=$ip; "Serial Number"=$serialNum; Reachable=$reachable}   
}

$entries | Sort-Object Reachable, "Last Checked" -Descending | Export-Csv -Path ".\Result.csv" -Encoding UTF8 -NoTypeInformation