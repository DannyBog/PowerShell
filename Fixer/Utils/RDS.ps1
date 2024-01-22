function Add-Option {
    $options = @()

    for ($i = 0; $i -lt $args.Count; $i++) {
        $options += [PSCustomObject]@{Option=$args[$i]}
    }

    return $options
}

$script:counter = 0
function Get-RDS {
    $users = (quser /server:$(-join $args[0..1]) 2>$null | Select-Object -Skip 1) -replace "\s\s+", ","
    if (!$users) {return}

    $rds = @()
    foreach ($user in $users) {
        if ($user.Split(",")[3] -eq "Active") {
            $script:counter++

            $server = $(-join $args[0..1])
            $samName = $user.Replace(" ", "").Split(",")[0]
            $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
            $sessionName = $user.Split(",")[1]
            $id = $user.Split(",")[2]
            $state = $user.Split(",")[3]
            $idleTime = $user.Split(",")[4]
            $logonTime = $user.Split(",")[5]
            $rds += [PSCustomObject]@{RDS=$server; User=($samName + " " + $displayName); SessionName=$sessionName; ID=$id; State=$state; "Idle Time"=$idleTime; "Logon Time"=$logonTime}   
        }

        if ($user.Split(",")[2] -eq "Disc" 2>$null) {
            $script:counter++

            $server = $(-join $args[0..1])
            $samName = $user.Replace(" ", "").Split(",")[0]
            $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
            $sessionName = ""
            $id = $user.Split(",")[1]
            $state = $user.Split(",")[2]
            $idleTime = $user.Split(",")[3]
            $logonTime = $user.Split(",")[4]
            $rds += [PSCustomObject]@{RDS=$server; User=($samName + " " + $displayName); SessionName=$sessionName; ID=$id; State=$state; "Idle Time"=$idleTime; "Logon Time"=$logonTime}
        }
    }

    return $rds
}

$name = Read-Host "Enter the naming convention of your terminal servers"
$num = Read-Host "Enter the number of terminal servers in your domain"
$entries = @()

for ($i = 1; $i -le $num; $i++) {
    Write-Progress -Activity "Gathering data..." -Status "$name$i" -PercentComplete ($i / 8 * 100);
    $entries += Get-RDS $name $i
}
Write-Progress -Activity "Completed" -Completed

while ($choice = $entries | Out-GridView –Title “RDS (User count: $script:counter)” -OutputMode Single) {
    $rds = $choice.RDS
    $user = $choice.User
    $id = $choice.ID
    $options = Add-Option "Shadow Session" "Kill Session"

    while ($choice = $options | Out-GridView -Title "$rds [$user]" -OutputMode Single) {
        switch ($choice.Option) {
            $options[0].Option {
                mstsc /v:$rds /shadow:$id /control
                break
            }

            $options[1].Option {
                reset session $id /server:$rds
                break
            }
        }
    }
}