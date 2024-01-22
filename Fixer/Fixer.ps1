function Add-Option {
    $options = @()

    for ($i = 0; $i -lt $args.Count; $i++) {
        $options += [PSCustomObject]@{Option=$args[$i]}
    }

    return $options
}

$options = Add-Option "Computers" "Printers" "RDS" "Subnets" "Users" "Groups"

while ($choice = $options | Out-GridView -Title "Fixer |Danny Bog| - 2.5v" -OutputMode Single) {
    Clear-Host

    switch ($choice.Option) {
        $options[0].Option {
            .\Utils\Computers.ps1
            break
        }

        $options[1].Option {
            .\Utils\Printers.ps1
            break
        }

        $options[2].Option {
            .\Utils\RDS.ps1
            break
        }

        $options[3].Option {
            .\Utils\Subnets.ps1
            break
        }

        $options[4].Option {
            .\Utils\Users.ps1
            break
        }

        $options[5].Option {
            .\Utils\Groups.ps1
            break
        }
    }
}