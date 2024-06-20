function Add-Option {
    $options = @()

    for ($i = 0; $i -lt $args.Count; $i++) {
        $options += [PSCustomObject]@{Option=$args[$i]}
    }

    return $options
}

function Get-Printers {
    $name = Read-Host "Enter the naming convention of your print servers"
    $num = Read-Host "Enter the number of print servers in your domain"
    $printers = @()

    for ($i = 1; $i -le $num; $i++) {
        $printers += (net view "\\$name$i" | Where-Object {$_ -match '\sPrint\s' }) -replace '\s\s+', ',' | ForEach-Object {[PSCustomObject]@{Server="$name$i"; Name=$_.Split(",")[0]; IP=$_.Split(",")[2]}}
    }

    return $printers
}

function Get-Toner {
    $toners = @()

    $config = Get-PrintConfiguration -ComputerName $args[1].Server -PrinterName $args[1].Name 2>$null

    if (!$config -or $config.Color) {
        switch ($args[0]) {
            {$_ -match "Epson" -or ($_ -match "Lexmark X792")} {
                $blackCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.0 -op:1.3.6.1.2.1.43.11.1.1.9.1.1 -q
                $cyanCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.1 -op:1.3.6.1.2.1.43.11.1.1.9.1.2 -q
                $magentaCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.2 -op:1.3.6.1.2.1.43.11.1.1.9.1.3 -q
                $yellowCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.3 -op:1.3.6.1.2.1.43.11.1.1.9.1.4 -q

                $blackTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.0 -op:1.3.6.1.2.1.43.11.1.1.8.1.1 -q
                $cyanTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.1 -op:1.3.6.1.2.1.43.11.1.1.8.1.2 -q
                $magentaTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.2 -op:1.3.6.1.2.1.43.11.1.1.8.1.3 -q
                $yellowTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.3 -op:1.3.6.1.2.1.43.11.1.1.8.1.4 -q

                break
            }

            {$_ -match "Lexmark CX725"} {
                $blackCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.1 -op:1.3.6.1.2.1.43.11.1.1.9.1.2 -q
                $cyanCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.3 -op:1.3.6.1.2.1.43.11.1.1.9.1.4 -q
                $magentaCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.6 -op:1.3.6.1.2.1.43.11.1.1.9.1.7 -q
                $yellowCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.8 -op:1.3.6.1.2.1.43.11.1.1.9.1.9 -q

                $blackTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.1 -op:1.3.6.1.2.1.43.11.1.1.8.1.2 -q
                $cyanTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.3 -op:1.3.6.1.2.1.43.11.1.1.8.1.4 -q
                $magentaTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.6 -op:1.3.6.1.2.1.43.11.1.1.8.1.7 -q
                $yellowTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.8 -op:1.3.6.1.2.1.43.11.1.1.8.1.9 -q

                break
            }

            {($_ -match "Samsung") -or ($_ -match "Lexmark C792") -or ($_ -match "ci")} {
                switch -Wildcard ($_) {
                    "Samsung CLP-775*" {
                        $blackCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.3 -op:1.3.6.1.2.1.43.11.1.1.9.1.4 -q
                        $cyanCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.2 -op:1.3.6.1.2.1.43.11.1.1.9.1.3 -q  
                        $magentaCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.1 -op:1.3.6.1.2.1.43.11.1.1.9.1.2 -q
                        $yellowCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.0 -op:1.3.6.1.2.1.43.11.1.1.9.1.1 -q

                        $blackTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.3 -op:1.3.6.1.2.1.43.11.1.1.8.1.4 -q
                        $cyanTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.2 -op:1.3.6.1.2.1.43.11.1.1.8.1.3 -q
                        $magentaTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.1 -op:1.3.6.1.2.1.43.11.1.1.8.1.2 -q
                        $yellowTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.0 -op:1.3.6.1.2.1.43.11.1.1.8.1.1 -q

                        break
                    }

                    default {
                        $blackCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.3 -op:1.3.6.1.2.1.43.11.1.1.9.1.4 -q
                        $cyanCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.0 -op:1.3.6.1.2.1.43.11.1.1.9.1.1 -q
                        $magentaCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.1 -op:1.3.6.1.2.1.43.11.1.1.9.1.2 -q
                        $yellowCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.2 -op:1.3.6.1.2.1.43.11.1.1.9.1.3 -q        

                        $blackTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.3 -op:1.3.6.1.2.1.43.11.1.1.8.1.4 -q
                        $cyanTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.0 -op:1.3.6.1.2.1.43.11.1.1.8.1.1 -q
                        $magentaTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.1 -op:1.3.6.1.2.1.43.11.1.1.8.1.2 -q
                        $yellowTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.2 -op:1.3.6.1.2.1.43.11.1.1.8.1.3 -q
                    }
                }
            }
        }

        if ($args[0] -match "Brother") {
            $toners += [PSCustomObject]@{Cartrige="Black"; Remaining="Not Supported"}
            $toners += [PSCustomObject]@{Cartrige="Cyan"; Remaining="Not Supported"}
            $toners += [PSCustomObject]@{Cartrige="Magenta"; Remaining="Not Supported"}
            $toners += [PSCustomObject]@{Cartrige="Yellow"; Remaining="Not Supported"}
        } else {
            $toners += [PSCustomObject]@{Cartrige="Black"; Remaining=[string](($blackCurrentLevel / $blackTotalLevel) * 100) + "%"}
            $toners += [PSCustomObject]@{Cartrige="Cyan"; Remaining=[string](($cyanCurrentLevel / $cyanTotalLevel) * 100) + "%"}
            $toners += [PSCustomObject]@{Cartrige="Magenta"; Remaining=[string](($magentaCurrentLevel / $magentaTotalLevel) * 100) + "%"}
            $toners += [PSCustomObject]@{Cartrige="Yellow"; Remaining=[string](($yellowCurrentLevel / $yellowTotalLevel) * 100) + "%"}
        }
    } else {
        switch ($args[0]) {
            {$_ -match "Lexmark"} {
                $blackCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.1 -op:1.3.6.1.2.1.43.11.1.1.9.1.2 -q
                $blackTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.1 -op:1.3.6.1.2.1.43.11.1.1.8.1.2 -q

                break
            }
            {($_ -match "Samsung") -or ($_ -match "Xerox") -or ($_ -match "4026iw") -or ($_ -match "6033DN")} {
                $blackCurrentLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.9.1.0 -op:1.3.6.1.2.1.43.11.1.1.9.1.1 -q
                $blackTotalLevel = .\Utils\snmpwalk -r:$args[1].IP -os:1.3.6.1.2.1.43.11.1.1.8.1.0 -op:1.3.6.1.2.1.43.11.1.1.8.1.1 -q

                break
            }
        }

        if ($args[0] -match "Brother") {
            $toners += [PSCustomObject]@{Cartrige="Black"; Remaining="Not Supported"}
        } else {
            $toners += [PSCustomObject]@{Cartrige="Black"; Remaining=[string](($blackCurrentLevel / $blackTotalLevel) * 100) + "%"}
        }
    }

    return $toners
}

while ($choice = Get-Printers | Sort-Object IP | Out-GridView –Title “Printers” -OutputMode Single) {
    $printer = $choice
    Set-Clipboard -Value $printer.IP

    if (Test-Connection $printer.IP -Count 2 2>$null) {
        $ip = "(" + $printer.IP + ")"
    } else {
        $ip = ""
    }

    $options = Add-Option "Check toner level" "Print test page"
    while ($choice = $options | Out-GridView -PassThru –Title "$($printer.Name)          $ip") {
        switch ($choice.Option) {
            {($_ -eq $options[0].Option) -and $ip} {
                $printerModel = .\Utils\snmpwalk -r:$printer.IP -os:1.3.6.1.2.1.25.3.2.1.3.0 -op:1.3.6.1.2.1.25.3.2.1.3.1 -q
                Get-Toner $printerModel $printer | Out-GridView –Title "$($printer.Name) | [$($printer.IP)] | {$printerModel}" -Wait
                break
            }

            {($_ -eq $options[1].Option) -and $ip} {
                rundll32.exe printui.dll, PrintUIEntry /k /n "\\$($printer.Server)\$($printer.Name)"
                break
            }
        }
    }
}