function Add-Option {
    $options = @()

    for ($i = 0; $i -lt $args.Count; $i++) {
        $options += [PSCustomObject]@{Option=$args[$i]}
    }

    return $options
}

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

function Parse-BuildNumber {
    $buildNumber = $args[0]

    switch ($buildNumber) {
        "3.10.511" {return "Windows NT 3.1"}
        "3.50.807" {return "Windows NT 3.5"}
        "3.10.528" {return "Windows NT 3.1, Service Pack 3"}
        "3.51.1057" {return "Windows NT 3.51"}
        "4.00.950" {return "Windows 95"}
        "4.00.950 A" {return "Windows 95 OEM Service Release 1"}
        "4.0.1381" {return "Windows NT 4.0"}
        "4.00.950 B" {return "Windows 95 OEM Service Release 2.1"}
        "4.00.950 C" {return "Windows 95 OEM Service Release 2.5"}
        "4.10.1998" {return "Windows 98"}
        "4.10.2222" {return "Windows 98 Econd Edition (SE)"}
        "5.0.2195" {return "Windows 2000"}
        "4.90.3000" {return "Windows Me"}
        "5.1.2600" {return "Windows XP"}
        "5.1.2600.1105-1106" {return "Windows XP, Service Pack 1"}
        "5.1.2600.2180" {return "Windows XP, Service Pack 2"}
        "6.0.6000" {return "Windows Vista"}
        "6.0.6001" {return "Windows Vista, Service Pack 1"}
        "5.1.2600" {return "Windows XP, Service Pack 3"}
        "6.0.6002" {return "Windows Vista, Service Pack 2"}
        "6.1.7600" {return "Windows 7"}
        "6.1.7601" {return "Windows 7, Service Pack 1"}
        "6.2.9200" {return "Windows 8"}
        "6.3.9200" {return "Windows 8.1"}
        "10.0.10240" {return "Windows 10, Version 1507"}
        "10.0.10586" {return "Windows 10, Version 1511"}
        "10.0.14393" {return "Windows 10, Version 1607"}
        "10.0.15063" {return "Windows 10, Version 1703"}
        "10.0.16299" {return "Windows 10, Version 1709"}
        "10.0.17134" {return "Windows 10, Version 1803"}
        "10.0.17763" {return "Windows 10, Version 1809"}
        "10.0.18362" {return "Windows 10, Version 1903"}
        "10.0.18363" {return "Windows 10, Version 1909"}
        "10.0.19041" {return "Windows 10, Version 2004"}
        "10.0.19042" {return "Windows 10, Version 20H2"}
        "10.0.19043" {return "Windows 10, Version 21H1"}
        "10.0.19044" {return "Windows 10, Version 21H2"}
        "10.0.19045" {return "Windows 10, Version 22H2"}
        "10.0.22000" {return "Windows 11, Version 21H2"}
        "10.0.22621" {return "Windows 11, Version 22H2"}
        "10.0.22631" {return "Windows 11, Version 23H2"}
    }
}

$computers = Get-ADComputer -Filter * -Properties Description, ms-Mcs-AdmPwd
$dn = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName
$dcs = nltest /dclist:$env:USERDNSDOMAIN | Select-Object -Skip 1
$dcs = $dcs | Select-String -Pattern "(?:^\s+)([a-zA-Z0-9-]+)(?=\.)" -AllMatches | ForEach-Object {$_.Matches.Groups[1].Value}

$entries = @()
$progress = 0

$entries += [PSCustomObject]@{Name="FileShareServer"; Username=""; "Last Checked"=""; IP="1.1.1.1"}

foreach ($computer in $computers) {
    Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($computers.Count) computers." -PercentComplete ($progress/$computers.Count * 100)
    $progress++

    $description = $computer.Description
    $name = $computer.Name
    $laps = $computer."ms-Mcs-AdmPwd"

    if (($description -match "User") -and ($description -match "Last checked") -and ($description -match "IP Address")) {
        $samName = $computer.Description.Split("()")[1]
        $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
        $lastChecked = ForEach-Object {Parse-Date $computer.Description.Split("()")[3]} 2>$null
        $ip = $computer.Description.Split("()")[5].SubString(1) -replace "\s+", ", "
        $ip = $ip.SubString(0, $ip.Length - 2)
    } else {
        $samName = ""
        $displayName = ""
        $lastChecked = ""
        $ip = ""
    }

    $entries += [PSCustomObject]@{Name=$name; User=($samName + " " + $displayName); "Last Checked"=$lastChecked; IP=$ip; LAPS=$laps}
}
Write-Progress -Activity "Completed" -Completed

$name = ""
$samName = ""
$displayName = ""
$entries = $entries | Sort-Object -Property @{e="Last Checked"; Descending=$true}, @{e="Name"; Descending=$false}

while ($choice = $entries | Out-GridView -Title "Computers" -OutputMode Single) {
    $computer = $choice | Select-Object Name, IP
    Set-Clipboard -Value $computer.Name

    $ip = (Test-Connection $computer.Name -Count 1 2>$null | Select-Object -ExpandProperty IPV4Address).IPAddressToString
    if ($ip) {
        $pc = $computer.Name
        $ip = "(" + $ip + ")"
    } else {
        $ip = Test-Connection $computer.IP -Count 1 2>$null | Select-Object -ExpandProperty Address
        if ($ip) {
            $pc = $ip
            $ip = "(" + $ip + ")"
        } else {
            $pc = ""
        }
    }

    if ($pc -ne "FileShareServer") {
        if ($dcs -contains $pc) {
            $options = Add-Option "Ping" "Remote Connect" "Tasks" "Services" "Registry" "System Info" "Group Policy" "Event Viewer" "Register DNS" "Restart PC"
        } else {
            $options = Add-Option "Ping" "Remote Connect" "Tasks" "Services" "Registry" "System Info" "Group Policy" "Disable Smart Card" "Register DNS" "Restart PC"
        }

        while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
            switch ($choice.Option) {
                $options[0].Option {
                    if ($pc) {
                        Start-Process powershell.exe -ArgumentList "ping $pc -t"
                    } else {
                        Start-Process powershell.exe -ArgumentList "ping $($computer.Name) -t"
                    }
                    
                    break
                }

                {($_ -eq $options[1].Option) -and $pc} {
                    $options = Add-Option "CmRc" "DameWare" "Shadow Session" "C$"
                    while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                        switch ($choice.Option) {
                            $options[0].Option {
                                .\Utils\CmRc\CmRcViewer.exe $pc
                                
                                break
                            }

                            $options[1].Option {
                                if (Test-Path "C:\Program Files\SolarWinds") {
                                    $dir = "C:\Program Files\SolarWinds"
                                } elseif (Test-Path "C:\Program Files (x86)\SolarWinds") {
                                    $dir =  "C:\Program Files (x86)\SolarWinds"
                                } else {
                                    $dir = ""
                                }

                                Push-Location $dir
                                $dameware = (Get-ChildItem -Path "DWRCC.exe" -Recurse).FullName
                                Pop-Location

                                if ($dameware) {
                                    & $dameware "-c:" "-h:" "-m:$pc" "-x:" "-a:1"
                                } else {
                                    Write-Host "DameWare is not installed on this machine."
                                }

                                break
                            }

                            $options[2].Option {
                                $entry = (quser /server:$pc 2>$null | Select-Object -Skip 1) -replace "\s\s+", ","
                                    
                                if ($entry) {
                                    $id = $entry.Split(",")[2]
                                    mstsc /v:$pc /shadow:$id /control $consent
                                }

                                break
                            }

                            $options[3].Option {
                                Invoke-Item "\\$pc\c$"
                                break
                            }
                        }
                        $options = Add-Option "CmRc" "DameWare" "Shadow Session" "C$"
                    }

                    break
                }

                {($_ -eq $options[2].Option) -and $pc} {
                    while ($choice = Get-Process -ComputerName $pc | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Multiple) {
                        $choice | ForEach-Object {taskkill /s $pc /IM $_.Id}
                    }

                    break
                }

                {($_ -eq $options[3].Option) -and $pc} {
                    while ($choice = Get-Service -ComputerName $pc | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Multiple) {
                        Restart-Service -InputObject $choice -Force
                    }

                    break
                }

                {($_ -eq $options[4].Option) -and $pc} {
                    $options = Add-Option "Adobe PDF Reader IE Add-on" "Unlimited Outlook search results (Cached mode only)" "RDP status" "Indexer backoff" "Require smard card" "Remove last logged on user" "Remove temporary user profiles" "Auto Accept (CmRcViewer)" "Auto Accept (DameWare)" "Auto Accept (Shadow Session)"
                    $extendedOptions = @()

                    $sids = reg query "\\$pc\HKU" 2>$null | Where-Object {($_ -like "*S-1-5-21*") -and (!$_.EndsWith("Classes"))}
                    $users = (quser /server:$pc 2>$null | Select-Object -Skip 1) -replace "\s\s+", ","

                    foreach ($sid in $sids) {
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid.Split("\")[1])
                        $objUser = try {$objSID.Translate([System.Security.Principal.NTAccount])} catch {}

                        $user = $users | Where-Object {($_.Split(",")[3] -eq "Active") -and ($objUser.Value -match $_.Replace(" ", "").Split(",")[0])}
                        if ($user) {break}
                   }

                    foreach ($option in $options) {
                        switch ($option.Option) {
                            {($_ -eq $options[0].Option) -and $user} {
                                $value = (reg query "\\$pc\$sid\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{CA8A9780-280D-11CF-A24D-444553540000}" /v "Flags" 2>$null) -replace "\s\s+", ","
                                if ($value) {
                                    $value = [int]$value.Split(",")[5]

                                    if ($value -eq 1) {
                                        $status = "Disabled"
                                    } else {
                                        $status = "Enabled"
                                    }
                                } else {
                                    $status = "Enabled"
                                }

                                break
                            }

                            {($_ -eq $options[1].Option) -and $user} {
                                $version = (reg query "\\$pc\HKLM\SOFTWARE\Classes\Outlook.Application\CurVer" /ve 2>$null) -replace "\s\s+", ","
                                if ($version) {$version = $version.Split(",")[5].Split(".")[2] + ".0"}
                                $value = (reg query "\\$pc\$sid\Software\Microsoft\Office\$version\Outlook\Search" /v "SearchResultsCap" 2>$null) -replace "\s\s+", ","

                                if ($value) {
                                    $value = [uint32]$value.Split(",")[5]

                                    if ($value -lt 4294967295) {
                                        $status = "Disabled"
                                    } else {
                                        $status = "Enabled"
                                    }
                                } else {
                                    $status = "Disabled"
                                }

                                break
                            }

                            $options[2].Option {
                                $value = (reg query "\\$pc\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" /v "fDenyTSConnections" 2>$null) -replace "\s\s+", ","
                                if ($value.Length -le 2) {
                                    $value = (reg query "\\$pc\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v "fDenyTSConnections" 2>$null) -replace "\s\s+", ","   
                                }

                                $value = [int]$value.Split(",")[5]
                                if ($value -eq 1) {
                                    $status = "Disabled"
                                } else {
                                    $status = "Enabled"
                                }

                                break
                            }

                            $options[3].Option {
                                $value = (reg query "\\$pc\HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "DisableBackoff" 2>$null) -replace "\s\s+", ","
                                if ($value) {
                                    $value = [int]$value.Split(",")[5]

                                    if ($value -eq 1) {
                                        $status = "Disabled"
                                    } else {
                                        $status = "Enabled"
                                    }
                                } else {
                                    $status = "Enabled"
                                }

                                break
                            }

                            $options[4].Option {
                                $value = (reg query "\\$pc\HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v "ExcludedCredentialProviders" 2>$null) -replace "\s\s+", ","
                                $value = $value.Split(",")[5]
                                if ($value) {
                                    $status = "Enabled"
                                } else {
                                    $status = "Disabled"
                                }

                                break
                            }

                            $options[5].Option {
                                $samName = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v "LastLoggedOnSAMUser" 2>$null) -replace "\s\s+", ","
                                $displayName = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v "LastLoggedOnDisplayName" 2>$null) -replace "\s\s+", ","
                                if (($samName -match "\w+") -and ($displayName -match "\w+")) {
                                    $samName = $samName.Split(",")[5].Split("\")[1]
                                    $displayName = "(" + $displayName.Split(",")[5] + ")"
                                    $status = $samName + " " + $displayName
                                } else {
                                    $status = ""
                                }

                                break
                            }

                            $options[6].Option {
                                $values = reg query "\\$pc\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" 2>$null
                                $profileRegs = @()
                                foreach ($value in $values) {
                                    if ($value.EndsWith(".bak")) {$profileRegs += $value}
                                }
                                
                                $profiles = @()
                                foreach ($profileReg in $profileRegs) {
                                    $sid = $profileReg.Split("\.")[6]
                                    $objUser = New-Object System.Security.Principal.SecurityIdentifier($sid)
                                    try {$profiles += $objUser.Translate([System.Security.Principal.NTAccount]).Value} catch {$profiles += $sid}
                                }

                                $status = $profiles

                                break
                            }

                            $options[7].Option {
                                $value1 = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "RemCtrl Taskbar Icon" 2>$null) -replace "\s\s+", ","
                                $value2 = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "Audible Signal" 2>$null) -replace "\s\s+", ","
                                $value3 = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "RemCtrl Connection Bar" 2>$null) -replace "\s\s+", ","
                                $value4 = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "Permission Required" 2>$null) -replace "\s\s+", ","

                                if ($value1) {$value1 = [int]$value1.Split(",")[5]}
                                if ($value2) {$value2 = [int]$value2.Split(",")[5]}
                                if ($value3) {$value3 = [int]$value3.Split(",")[5]}
                                if ($value4) {$value4 = [int]$value4.Split(",")[5]}

                                $value = $value1 + $value2 + $value3 + $value4

                                if ($value -eq 0) {
                                    $status = "Enabled"
                                } elseif ($value4 -eq 0) {
                                    $status = "Partially Enabled"
                                } else {
                                    $status = "Disabled"
                                }

                                break
                            }

                            $options[8].Option {
                                $value1 = (reg query "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "TrayIcon" 2>$null) -replace "\s\s+", ","
                                $value2 = (reg query "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "No Notify Sound" 2>$null) -replace "\s\s+", ","
                                $value3 = (reg query "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "Notify On New Connection" 2>$null) -replace "\s\s+", ","
                                $value4 = (reg query "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "Permission Required" 2>$null) -replace "\s\s+", ","

                                if ($value1) {$value1 = [int]$value1.Split(",")[5]}
                                if ($value2) {$value2 = [int]$value2.Split(",")[5]}
                                if ($value3) {$value3 = [int]$value3.Split(",")[5]}
                                if ($value4) {$value4 = [int]$value4.Split(",")[5]}

                                $value = $value1 + $value2 + $value3 + $value4

                                if ($value -eq 0) {
                                    $status = "Enabled"
                                } elseif ($value4 -eq 0) {
                                    $status = "Partially Enabled"
                                } else {
                                    $status = "Disabled"
                                }

                                break
                            }

                            $options[9].Option {
                                $value = (reg query "\\$pc\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" /v "Shadow" 2>$null) -replace "\s\s+", ","

                                if ($value) {
                                    $value = [int]$value.Split(",")[5]

                                    if ($value -eq 2) {
                                        $consent = ""
                                        $status = "Enabled"
                                    } else {
                                        $status = "Disabled"
                                    }
                                } else {
                                    $status = "Disabled"
                                }

                                break
                            }

                            default {
                                $status = "Unknown"
                                break
                            }
                        }

                        $extendedOptions += [PSCustomObject]@{Option=$option | Select-Object -ExpandProperty Option; Status=$status}
                    }

                    while ($choice = $extendedOptions | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                        switch ($choice.Option) {
                            {($_ -eq $extendedOptions[0].Option) -and $user} {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[0].Status = "Disabled"
                                    $value = 1
                                } else {
                                    $extendedOptions[0].Status = "Enabled"
                                    $value = 0
                                }
                                
                                reg add "\\$pc\$sid\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{CA8A9780-280D-11CF-A24D-444553540000}" /v "Flags" /t REG_DWORD /d $value /f | Out-Null

                                break
                            }

                            {($_ -eq $extendedOptions[1].Option) -and $user -and $version} {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[1].Status = "Disabled"
                                    reg delete "\\$pc\$sid\Software\Microsoft\Office\$version\Outlook\Search" /v "SearchResultsCap" /f | Out-Null
                                } else {
                                    $extendedOptions[1].Status = "Enabled"
                                    reg add "\\$pc\$sid\Software\Microsoft\Office\$version\Outlook\Search" /v "SearchResultsCap" /t REG_DWORD /d 4294967295 /f | Out-Null
                                }

                                break
                            }

                            $extendedOptions[2].Option {
                                if ($choice.Status -eq "Disabled") {
                                    $extendedOptions[2].Status = "Enabled"
                                    reg add "\\$pc\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v "fDenyTSConnections" /t REG_DWORD /d 0 /f | Out-Null
                                }

                                break
                            }

                            $extendedOptions[3].Option {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[3].Status = "Disabled"
                                    $value = 1
                                } else {
                                    $extendedOptions[3].Status = "Enabled"
                                    $value = 0
                                }
                                
                                reg add "\\$pc\HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v "DisableBackoff" /t REG_DWORD /d $value /f | Out-Null
                                Get-Service -Name "WSearch" -ComputerName $pc | Restart-Service -Force

                                break
                            }

                            $extendedOptions[4].Option {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[4].Status = "Disabled"
                                    reg delete "\\$pc\HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v "ExcludedCredentialProviders" /f | Out-Null
                                } else {
                                    $extendedOptions[4].Status = "Enabled"
                                    reg add "\\$pc\HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v "ExcludedCredentialProviders" /t REG_SZ /d "{60b78e88-ead8-445c-9cfd-0b87f74ea6cd}" /f | Out-Null
                                }

                                $id = (Get-Process -Name "LogonUI" -ComputerName $pc).Id 2>$null
                                if ($id) {taskkill /s $pc /IM $id 1>$null}

                                break
                            }

                            $extendedOptions[5].Option {
                                if ($choice.Status) {
                                    $extendedOptions[5].Status = ""
                                    reg delete "\\$pc\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /va /f | Out-Null
                                }

                                break
                            }

                            $extendedOptions[6].Option {
                                if ($choice.Status) {
                                    $extendedOptions[6].Status = ""
                                    foreach ($profileReg in $profileRegs) {
                                        reg delete \\$pc\$profileReg /f | Out-Null
                                    }
                                }

                                break
                            }

                            $extendedOptions[7].Option {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[7].Status = "Disabled"
                                    $value = 1
                                } else {
                                    $extendedOptions[7].Status = "Enabled"
                                    $value = 0
                                }

                                reg add "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "RemCtrl Taskbar Icon" /t REG_DWORD /d $value /f | Out-Null
                                reg add "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "Audible Signal" /t REG_DWORD /d $value /f | Out-Null
                                reg add "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "RemCtrl Connection Bar" /t REG_DWORD /d $value /f | Out-Null
                                reg add "\\$pc\HKLM\SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control" /v "Permission Required" /t REG_DWORD /d $value /f | Out-Null

                                break
                            }

                            $extendedOptions[8].Option {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[8].Status = "Disabled"
                                    $value = 1
                                } else {
                                    $extendedOptions[8].Status = "Enabled"
                                    $value = 0
                                }

                                reg add "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "TrayIcon" /t REG_DWORD /d $value /f | Out-Null
                                reg add "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "No Notify Sound" /t REG_DWORD /d $value /f | Out-Null
                                reg add "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "Notify On New Connection" /t REG_DWORD /d $value /f | Out-Null
                                reg add "\\$pc\HKLM\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" /v "Permission Required" /t REG_DWORD /d $value /f | Out-Null

                                break
                            }

                            $extendedOptions[9].Option {
                                if ($choice.Status -eq "Enabled") {
                                    $extendedOptions[9].Status = "Disabled"
                                    $consent = ""
                                    $value = 1
                                } else {
                                    $extendedOptions[9].Status = "Enabled"
                                    $consent = "/noConsentPrompt"
                                    $value = 2
                                }
                                
                                reg add "\\$pc\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" /v "Shadow" /t REG_DWORD /d $value /f | Out-Null

                                break
                            }
                        }
                    }

                    break
                }

                {($_ -eq $options[5].Option) -and $pc} {
                    $options = Add-Option "Network Info" "Hardware Info" "Software Info" "Groups Info"
                    while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                        switch ($choice.Option) {
                            $options[0].Option {
                                $subnet = Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" | Select-Object -ExpandProperty IPSubnet
                                $gateway = Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" | Select-Object -ExpandProperty DefaultIPGateway

                                $dc = nltest /server:$pc /dsgetdc:$env:USERDOMAIN
                                $dcName = $dc.Split(":")[1].SubString(3)
                                $dcIP = $dc.Split(":")[3].SubString(3)
                                $dc = $dcName + " (" + $dcIP + ")"

                                $dhcp = (Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" | Where-Object {$_.IPAddress -match $ip}).DHCPEnabled
                                $dhcpServer = Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration | Select-Object -ExpandProperty DHCPServer
                                if ($dhcp) {
                                    $dhcpServer = ((Resolve-DnsName $dhcpServer).NameHost.Split(".")[0].ToUpper()) + " (" + $dhcpServer + ")"
                                } else {
                                    $dhcpServer = "-"
                                }
                                
                                $dnsServers = Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration | Select-Object -ExpandProperty DNSServerSearchOrder
                                foreach ($dnsServer in $dnsServers) {
                                    $servers += ((Resolve-DnsName $dnsServer).NameHost.Split(".")[0].ToUpper()) + " (" + $dnsServer + ")" + ", "
                                }
                                $dnsServers = $servers.TrimEnd(", ")
                                $servers = ""

                                $specs = [ordered]@{
                                    "Subnet Mask" = $subnet
                                    "Default Gateway" = $gateway
                                    "DC" = $dc
                                    "DHCP Enabled" = $dhcp
                                    "DHCP Server" = $dhcpServer
                                    "DNS Servers" = $dnsServers
                                }
                        
                                $result = $specs | Out-GridView -Title "$($computer.Name)     $ip" -PassThru
                                if ($result) {$result | Export-Csv -Path ".\$($computer.Name).csv" -Encoding UTF8 -NoTypeInformation}
                            }

                            $options[1].Option {
                                $vendor = (Get-WmiObject -ComputerName $pc -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty Vendor).Split(" ")[0]

                                switch ($vendor) {
                                    {@("Dell", "HP") -contains $_} {
                                        $model = Get-WmiObject -ComputerName $pc -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty Name
                                    }

                                    "Lenovo" {
                                        $vendor = (Get-Culture).TextInfo.ToTitleCase($vendor.ToLower())
                                        $model = Get-WmiObject -ComputerName $pc -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty Version
                                        
                                    }

                                    default {
                                        $vendor = "?"
                                        $model = "?"
                                        $serialNum = "?"
                                    }
                                }

                                $serialNum = Get-WmiObject -ComputerName $pc -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty IdentifyingNumber
                                $mac = (Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration | Where-Object {$_.IPAddress -match $ip}).MACAddress
                                $ram = (Get-WmiObject -ComputerName $pc -Class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB
                                $partitionTable = (Get-WmiObject -ComputerName $pc -Class Win32_DiskPartition -Filter "Index=0" | Select-Object -ExpandProperty Type).Split(":")[0]
                                if ($partitionTable -ne "GPT") {$partitionTable = "MBR"}
                                $disk = Get-WmiObject -ComputerName $pc -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
                                $diskSize = "{0:n2}" -f (($disk | Measure-Object -Property Size -Sum).Sum / 1GB)
                                $diskSpace = "{0:n2}" -f (($disk | Measure-Object -Property FreeSpace -Sum).Sum / 1GB)
                                $processor = Get-WmiObject -ComputerName $pc -Class Win32_Processor | Select-Object -ExpandProperty Name

                                $build = Get-WmiObject -ComputerName $pc -Class Win32_OperatingSystem | Select-Object -ExpandProperty Version
                                $language = Get-WmiObject -ComputerName $pc -Class Win32_OperatingSystem | Select-Object -ExpandProperty MUILanguages | Select-Object -First 1
                                $osType = Get-WmiObject -ComputerName $pc -Class Win32_OperatingSystem | Select-Object -ExpandProperty ProductType
                                switch ($osType) {
                                    1 {
                                        $value = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2>$null) -replace "\s\s+", ","
                                        $UBR = [int]$value.Split(",")[5]

                                        $os = Parse-BuildNumber $build
                                        $os += " (" + $build + "." + $UBR + ")"
                                        $os += " [" + $language + "]"
                                        break
                                    }

                                    {2 -or 3} {
                                        $value = (reg query "\\$pc\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2>$null) -replace "\s\s+", ","
                                        $UBR = [int]$value.Split(",")[5]

                                        $os = Get-WmiObject -ComputerName $pc -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
                                        $os += " (" + $build + "." + $UBR + ")"
                                        $os += " [" + $language + "]"
                                        break
                                    }
                                }
                                
                                $lastReboot = (Get-WmiObject -ComputerName $pc -Class Win32_OperatingSystem | Select-Object @{n='LastBootUpTime'; e={$_.ConvertToDateTime($_.LastBootUpTime)}} | Select-Object -ExpandProperty LastBootUpTime).ToString("dd/MM/yyyy HH:mm:ss")

                                $specs = [ordered]@{
                                    "Vendor" = $vendor
                                    "Model" = $model
                                    "Serial Number" = $serialNum
                                    "MAC Address" = $mac
                                    "RAM (GB)" = $ram
                                    "Partition Table" = $partitionTable
                                    "Disk Size (GB)" = $diskSize
                                    "Disk Space (GB)" = $diskSpace
                                    "Processor" = $processor
                                    "Operating System" = $os
                                    "Last Reboot" = $lastReboot
                                }
                        
                                $result = $specs | Out-GridView -Title "$($computer.Name)     $ip" -PassThru
                                if ($result) {$result | Export-Csv -Path ".\$($computer.Name).csv" -Encoding UTF8 -NoTypeInformation}

                                break
                            }

                            $options[2].Option {
                                Write-Progress -Activity "Gathering Data..."
                                $machineKeys32 = reg query "\\$pc\HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 2>$null | Where-Object {$_ -ne ""}
                                $machineKeys64 = reg query "\\$pc\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" 2>$null | Where-Object {$_ -ne ""}

                                $sids = reg query "\\$pc\HKU" 2>$null | Where-Object {($_ -match "S-1-5-21") -and (!$_.EndsWith("Classes"))}
                                $users = (quser /server:$pc 2>$null | Select-Object -Skip 1) -replace "\s\s+", ","

                                foreach ($sid in $sids) {
                                    $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid.Split("\")[1])
                                    $objUser = try {$objSID.Translate([System.Security.Principal.NTAccount])} catch {}
    
                                    $user = $users | Where-Object {($_.Split(",")[3] -eq "Active") -and ($objUser.Value -match $_.Replace(" ", "").Split(",")[0])}
                                    if ($user) {
                                        $userKeys32 = reg query "\\$pc\$sid\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 2>$null | Where-Object {$_ -ne ""}
                                        $userKeys64 = reg query "\\$pc\$sid\Software\Microsoft\Windows\CurrentVersion\Uninstall" 2>$null | Where-Object {$_ -ne ""}
                                        $sid = ""

                                        break
                                    }
                                }

                                $keys = $machineKeys32 + $machineKeys64 + $userKeys32 + $userKeys64

                                [System.Collections.ArrayList]$programs = @()
                                foreach ($key in $keys) {
                                    $key = "\\$pc\" + $key

                                    $displayNameValue = (reg query $key 2>$null | Where-Object {($_ -match "DisplayName")}) -replace "\s\s+", ","
                                    $systemComponentValue = (reg query $key 2>$null | Where-Object {($_ -match "SystemComponent")}) -replace "\s\s+", ","
                                    $parentKeyNameValue = (reg query $key 2>$null | Where-Object {($_ -match "ParentKeyName")}) -replace "\s\s+", ","
                                    $uninstallStringValue = (reg query $key 2>$null | Where-Object {($_ -match "UninstallString")}) -replace "\s\s+", ","

                                    if ($displayNameValue -and !$systemComponentValue -and !$parentKeyNameValue -and $uninstallStringValue) {
                                        $publisherValue = (reg query $key 2>$null | Where-Object {($_ -match "Publisher")}) -replace "\s\s+", ","
                                        $installedOnValue = (reg query $key 2>$null | Where-Object {($_ -match "InstallDate")}) -replace "\s\s+", ","
                                        $sizeValue = (reg query $key 2>$null | Where-Object {($_ -match "Size")}) -replace "\s\s+", ","
                                        $versionValue = (reg query $key 2>$null | Where-Object {($_ -match "DisplayVersion")}) -replace "\s\s+", ","

                                        $name = $displayNameValue.Split(",")[3]
                                        if ($publisherValue) {
                                            $publisher =  $publisherValue.Split(",")[3]
                                        } else {
                                            $publisher = ""
                                        }

                                        if ($installedOnValue) {
                                            $unformattedDate = $installedOnValue.Split(",")[3]
                                            $installedOn = $unformattedDate.SubString(6) + "/" + $unformattedDate.SubString(4, 2) + "/" + $unformattedDate.SubString(0, 4)
                                        } else {
                                            $installedOn = ""
                                        }

                                        if ($versionValue) {
                                            $version = $versionValue.Split(",")[3]
                                        } else {
                                            $version = ""
                                        }

                                        $uninstallString = $uninstallStringValue.Split(",")[3]
                                        if ($uninstallString.StartsWith("MsiExec", "CurrentCultureIgnoreCase")) {
                                            $uninstallable = "✓"
                                        } else {
                                            $uninstallable = ""
                                        }

                                        $programs.Add([PSCustomObject]@{Name=$name; Publisher=$publisher; "Installed On"=$installedOn; Version=$version; "Uninstallable"=$uninstallable; UninstallString=$uninstallString}) | Out-Null
                                    }
                                }
                                Write-Progress -Activity "Completed" -Completed

                                while ($choice = $programs | Select-Object -Property * -ExcludeProperty UninstallString | Sort-Object Name | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                                    $program = $programs | Where-Object {$_.Name -eq $choice.Name}
                                    if ($program.uninstallable) {
                                        $uninstall = $program.UninstallString + " " + "/passive" + " " + "/qn"
                                        cmd /c ".\Utils\PsExec.exe -s -d -accepteula -nobanner \\$pc $uninstall" 2>$null
                                        $programs.Remove($program)
                                    }
                                }

                                break
                            }

                            $options[3].Option {
                                $options = Add-Option "Local Administrators" "Allowed RDP Users" "Allowed SCCM Users"
                                while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                                    switch ($choice.Option) {
                                        $options[0].Option {
                                            $users = Get-WmiObject -ComputerName $pc -Class Win32_GroupUser | Where-Object {$_.GroupComponent -like '*"Administrators"'}
                                            break
                                        }

                                        $options[1].Option {
                                            $users = Get-WmiObject -ComputerName $pc -Class Win32_GroupUser | Where-Object {$_.GroupComponent -like '*"Remote Desktop Users"'}
                                            break
                                        }

                                        $options[2].Option {
                                            $users = Get-WmiObject -ComputerName $pc -Class Win32_GroupUser | Where-Object {$_.GroupComponent -like '*"ConfigMgr Remote Control Users"'}
                                            break
                                        }
                                    }

                                    $members = @()
                                    foreach ($user in $users) {
                                        $members += $user.PartComponent.Split('""')[1] + "\" + $user.PartComponent.Split('""')[3]
                                    }

                                    $result = $members | Select-Object @{n="Member"; e={$_}} | Out-GridView -Title "$($computer.Name)     $ip" -PassThru
                                    if ($result) {$result | Export-Csv -Path ".\$($computer.Name).csv" -Encoding UTF8 -NoTypeInformation}
                                }

                                break
                            }
                        }
                        $options = Add-Option "Network Info" "Hardware Info" "Software Info" "Groups Info"
                    }

                    break
                }

                
                {($_ -eq $options[6].Option) -and $pc} {
                    $options = Add-Option "GP Result" "GP Update"

                    while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                        switch ($choice.Option) {
                            $options[0].Option {
                                $users = (quser /server:$pc 2>$null | Select-Object -Skip 1) -replace "\s\s+", ","
                                $lockedUsers = $users | Where-Object {($_.Split(",")[2] -eq "Disc")}
                                $maybeLockedUser = $users | Where-Object {($_.Split(",")[3] -eq "Active") -and ($_.Split(",")[1] -notlike "rdp-tcp*")}
                                $rdpUsers = $users | Where-Object {($_.Split(",")[1] -like "rdp-tcp*")}

                                $usersList = @()

                                if ($maybeLockedUser) {
                                    $samName = $maybeLockedUser.Replace(" ", "").Split(",")[0]
                                    $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
                                    $maybeLockedUser = $samName + " " + $displayName

                                    $processes = Get-Process "LogonUI" -ComputerName $pc 2>$null
                                    if ($users.Count -eq $processes.Count) {
                                        $usersList += [PSCustomObject]@{User=$maybeLockedUser; Status="Locked"}
                                    } else {
                                        $usersList += [PSCustomObject]@{User=$maybeLockedUser; Status="Active"}
                                    }
                                }

                                foreach ($lockedUser in $lockedUsers) {
                                    $samName = $lockedUser.Replace(" ", "").Split(",")[0]
                                    $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
                                    $lockedUser = $samName + " " + $displayName
                                    $usersList += [PSCustomObject]@{User=$lockedUser; Status="Locked"}
                                }
    
                                foreach ($rdpUser in $rdpUsers) {
                                    $samName = $rdpUser.Replace(" ", "").Split(",")[0]
                                    $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
                                    $rdpUser = $samName + " " + $displayName
                                    $usersList += [PSCustomObject]@{User=$rdpUser; Status="RDP"}
                                }

                                $user = $usersList | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single
                                
                                if ($user) {
                                    $user = $user.User.Split(" ")[0]
                                    $gpresult = gpresult /s $pc /user $user /r
                                } elseif (!$usersList) {
                                    $gpresult = gpresult /s $pc /scope computer /r
                                } else {
                                    break
                                }

                                $computerPolicies = @()
                                $userPolicies = @()
                                $computerSettings = $false
                                $userSettings = $true
                                $flag = $false
                                foreach ($line in $gpresult) {
                                    if (!$line) {continue}

                                    if ($line -match "Applied Group Policy Objects") {
                                        $computerSettings = !$computerSettings
                                        $userSettings = !$userSettings
                                        $flag = $true
                                        continue
                                    }

                                    if ($line -match "The following GPOs were not applied because they were filtered out") {
                                        $flag = $false
                                        continue
                                    }

                                    if ($flag) {
                                        if ($computerSettings) {$computerPolicies += $line}
                                        if ($userSettings) {$userPolicies += $line}
                                    }
                                }

                                $computerPolicies = ($computerPolicies | Select-Object -Skip 1 -Unique) -replace "\s\s+", ""
                                $userPolicies = ($userPolicies | Select-Object -Skip 1 -Unique) -replace "\s\s+", ""
                                $policies = @()

                                foreach ($policy in $computerPolicies) {
                                    $policies += [PSCustomObject]@{Configuration="Computer"; Policy=$policy}
                                }

                                foreach ($policy in $userPolicies) {
                                    $policies += [PSCustomObject]@{Configuration="User"; Policy=$policy}
                                }

                                $result = $policies | Sort-Object Configuration, Policy | Out-GridView -Title "$($computer.Name)" -PassThru
                                if ($result) {$result | Export-Csv -Path ".\$($computer.Name).csv" -Encoding UTF8 -NoTypeInformation}

                                break
                            }

                            $options[1].Option {
                                Invoke-GPUpdate -Computer $pc -RandomDelayInMinutes 0 -Force
                                break
                            }
                        }
                    }

                    break
                }

                {($_ -eq $options[7].Option) -and $pc} {
                    if ($dcs -contains $pc) {
                        $events = @("4722", "4723", "4724", "4725", "4728", "4729", "4737", "4738", "4740", "4750", "5141")

                        $descriptions = @(
                            "A user account was enabled"
                            "An attempt was made to change an account's password"
                            "An attempt was made to reset an accounts password"
                            "A user account was disabled"
                            "A member was added to a security-enabled global group"
                            "A member was removed from a security-enabled global group"
                            "A security-enabled global group was changed"
                            "A user account was changed"
                            "A user account was locked out"
                            "A security-disabled global group was changed"
                            "A directory service object was deleted"
                        )

                        $maxLength = [Math]::Max($events.Length, $descriptions.Length)

                        $options = @()
                        for ($i = 0; $i -lt $maxLength; $i++) {
                            $options += [PSCustomObject]@{Option=$events[$i]; Description=$descriptions[$i]}
                        }

                        while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                            switch ($choice.Option) {
                                {$_ -in "4728", "4729"} {
                                    Write-Progress -Activity "Gathering Data..."
                                    $events = Get-WinEvent -ComputerName $pc -FilterHashtable @{LogName="Security"; ID=$($choice.Option)} 2>$null

                                    $parsedEvents = @()

                                    foreach ($event in $events) {
                                        $distinguishedName = $event.Properties[0].Value
                                        $displayName1 = &{if ($name = (Get-ADObject -Filter {DistinguishedName -eq $distinguishedName}).Name) {$name} else {($distinguishedName.Split(",") | ConvertFrom-StringData).CN + " " + "(" + "?" + ")"}}
                                        $samName = $event.Properties[6].Value
                                        $displayName2 = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
	                                    $group = $event.Properties[2].Value
	                                    $timeStamp = $event.TimeCreated

	                                    $parsedEvents += [PSCustomObject]@{User=($samName + " " + $displayName2); Member=$displayName1; Group=$group; Timestamp=$timeStamp}
                                    }
                                    Write-Progress -Activity "Completed" -Completed

                                    $parsedEvents | Out-GridView -Title "$($computer.Name)     $ip" -Wait

                                    break
                                }

                                "4738" {
                                    Write-Progress -Activity "Gathering Data..."
                                    $events = Get-WinEvent -ComputerName $pc -FilterHashtable @{LogName="Security"; ID=$($choice.Option)} 2>$null

                                    $parsedEvents = @()

                                    foreach ($event in $events) {
                                        $samName = $event.Properties[1].Value
                                        $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
	                                    $source = $event.Properties[5].Value
	                                    $timeStamp = $event.TimeCreated

	                                    $parsedEvents += [PSCustomObject]@{User=($samName + " " + $displayName); Source=$source; Timestamp=$timeStamp}
                                    }
                                    Write-Progress -Activity "Completed" -Completed

                                    $parsedEvents | Out-GridView -Title "$($computer.Name)     $ip" -Wait

                                    break
                                }

                                "4740" {
                                    Write-Progress -Activity "Gathering Data..."
                                    $events = Get-WinEvent -ComputerName $pc -FilterHashtable @{LogName="Security"; ID=$($choice.Option)} 2>$null

                                    $parsedEvents = @()

                                    foreach ($event in $events) {
                                        $samName = $event.Properties[0].Value
                                        $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
	                                    $source = $event.Properties[1].Value
	                                    $timeStamp = $event.TimeCreated

	                                    $parsedEvents += [PSCustomObject]@{User=($samName + " " + $displayName); Source=$source; Timestamp=$timeStamp}
                                    }
                                    Write-Progress -Activity "Completed" -Completed

                                    $parsedEvents | Out-GridView -Title "$($computer.Name)     $ip" -Wait

                                    break
                                }

                                "5141" {
                                    Write-Progress -Activity "Gathering Data..."
                                    $events = Get-WinEvent -ComputerName $pc -FilterHashtable @{LogName="Security"; ID=$($choice.Option)} 2>$null

                                    $parsedEvents = @()

                                    foreach ($event in $events) {
                                        $samName = $event.Properties[3].Value
                                        $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
                                        $class = $event.Properties[10].Value
	                                    $object = $event.Properties[8].Value
                                        if ($object) {$object = ($object.Split(",") | ConvertFrom-StringData).CN | Select-Object -First 1}
	                                    $timeStamp = $event.TimeCreated

	                                    $parsedEvents += [PSCustomObject]@{User=($samName + " " + $displayName); Class=$class; Object=$object; Timestamp=$timeStamp}
                                    }
                                    Write-Progress -Activity "Completed" -Completed

                                    $parsedEvents | Out-GridView -Title "$($computer.Name)     $ip" -Wait

                                    break
                                }

                                default {
                                    Write-Progress -Activity "Gathering Data..."
                                    $events = Get-WinEvent -ComputerName $pc -FilterHashtable @{LogName="Security"; ID=$($choice.Option)} 2>$null

                                    $parsedEvents = @()

                                    foreach ($event in $events) {
                                        $samName = $event.Properties[0].Value
                                        $displayName1 = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
	                                    $source = $event.Properties[4].Value
                                        $displayName2 = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $source}).Name) {$name} else {"?"}}) + ")"
	                                    $timeStamp = $event.TimeCreated

	                                    $parsedEvents += [PSCustomObject]@{User=($samName + " " + $displayName1); Source=($source + " " + $displayName2); Timestamp=$timeStamp}
                                    }
                                    Write-Progress -Activity "Completed" -Completed

                                    $parsedEvents | Out-GridView -Title "$($computer.Name)     $ip" -Wait

                                    break
                                }
                            }
                        }
                    } else {
                        $options = Add-Option "One Day" "One Week"
                        while ($choice = $options | Out-GridView -Title "$($computer.Name)     $ip" -OutputMode Single) {
                            switch ($choice.Option) {
                                $options[0].Option {
                                    Get-ADComputer -Identity $computer.Name | Move-ADObject -TargetPath "OU=DisabledSmartCard1Day,OU=Workstations,$dn"
                                    break
                                }

                                $options[1].Option {
                                    Get-ADComputer -Identity $computer.Name | Move-ADObject -TargetPath "OU=DisableSmartCard1Week,OU=Workstations,$dn"
                                    break
                                }
                            }
                            $options = Add-Option "One Day" "One Week"
                        }
                    }

                    break
                }

                {($_ -eq $options[8].Option) -and $pc} {
                    $adapter = Get-WmiObject -ComputerName $pc -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" | Where-Object {$_.IPAddress -match $ip}
                    $adapter.SetDynamicDNSRegistration($true, $false) | Out-Null
                    ([WMIClass]"\\$pc\ROOT\CImv2:Win32_Process").Create("cmd.exe /c ipconfig /registerdns") | Out-Null
                    
                    break
                }

                {($_ -eq $options[9].Option) -and $pc} {
                    Restart-Computer -ComputerName $pc -Force
                    break
                }
            }

            if ($dcs -contains $pc) {
                $options = Add-Option "Ping" "Remote Connect" "Tasks" "Services" "Registry" "System Info" "Group Policy" "Event Viewer" "Register DNS" "Restart PC"
            } else {
                $options = Add-Option "Ping" "Remote Connect" "Tasks" "Services" "Registry" "System Info" "Group Policy" "Disable Smart Card" "Register DNS" "Restart PC"
            }
        }
    } else {
        $openFiles = (openfiles /query /s $pc /fo CSV /nh /v) | ForEach-Object {$_.Replace("`"", "")} | Where-Object {[System.IO.Path]::GetExtension($_)}

        [System.Collections.ArrayList]$files = @()
        $progress = 0

        foreach ($openFile in $openFiles) {
            Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($openFiles.Length) files." -PercentComplete ($progress/$openFiles.Length * 100)
            $progress++

            $id = $openFile.Split(",")[1]
            $samName = $openFile.Split(",")[2]
            $accessedBy = $samName + " " + "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
            $type = $openFile.Split(",")[3]
            $numLocks = $openFile.Split(",")[4]
            $openMode = $openFile.Split(",")[5]

            $files += [PSCustomObject]@{"Open File"=$file; ID=$id; "Accessed By"=$accessedBy; Type=$type; "# Locks"=$numLocks; "Open Mode"=$openMode}
        }
        Write-Progress -Activity "Completed" -Completed

        while ($choices = $files | Sort-Object "Accessed By" | Out-GridView -Title "$($computer.Name)     $ip" -PassThru) {
            foreach ($choice in $choices) {
                $file = $files | Where-Object {$_.ID -eq $choice.ID}
                openfiles /disconnect /s $pc /id $file.ID
                $files.Remove($file)
            }
        }
    }
}