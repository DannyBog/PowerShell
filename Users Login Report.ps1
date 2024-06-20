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

$users = Get-ADUser -Filter {Enabled -eq "true"} -Properties Description, whenCreated, whenChanged | Select-Object Name, SamAccountName, Description, @{n="Created"; e={$_.whenCreated}}, @{n="Modified"; e={$_.whenChanged}} | Sort-Object Name

$entries = @()
$progress = 0

foreach ($user in $users) {
    Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($users.Count) users." -PercentComplete ($progress/$users.Count * 100)
    $progress++

    $computers = @()
    Get-ADComputer -Filter * -Properties Description | ForEach-Object {if ($_.Description -and $_.Description.Split("()")[1] -eq $user.SamAccountName) {$computers += $_}}
    foreach ($computer in $computers) {
        $samName = $computer.Description.Split("()")[1]
        $lastChecked = ForEach-Object {Parse-Date $computer.Description.Split("()")[3]} 2>$null
        $ip = $computer.Description.Split("()")[5]

        $entries += [PSCustomObject]@{Name=$user.Name; Description=$user.Description; Created=$user.Created; Modified=$user.Modified; "Last Accessed Workstation"=$computer.Name; IP=$ip; "Last Checked"=$lastChecked}
    }

    if (!$computers) {$entries += [PSCustomObject]@{Name=$user.Name; Description=$user.Description; Created=$user.Created; Modified=$user.Modified; "Last Accessed Workstation"=""; IP=""; "Last Checked"=""}}
}
Write-Progress -Activity "Completed" -Completed

if ($entries) {$entries | Export-Csv -Path ".\Report.csv" -Encoding UTF8 -NoTypeInformation}