$global:success = $false

function Print-Sudoku {
    param (
        [int[]]$board
    )

    for ($i = 0; $i -lt $board.Count; $i++) {
        if (($i + 1) % 9 -eq 0) {
            Write-Host $board[$i]
        } else {
            Write-Host -NoNewline "$($board[$i]), "
        }
    }
}

function Check-Row {
    param (
        [int[]]$board,
        [int]$pos,
        [int]$num
    )

    $col = $pos % 9
    for ($row = 0; $row -lt 9; $row++) {
        if ($board[$col + $row * 9] -eq $num) {return $false}
    }

    return $true
}

function Check-Column {
    param (
        [int[]]$board,
        [int]$pos,
        [int]$num
    )

    $row = [Math]::Truncate($pos / 9) # Casting to [int] invariably employs half-to-even rounding, where numbers with a fractional part of .5 are rounded to the nearest even integer (whether positive or negative).
    for ($col = 0; $col -lt 9; $col++) {
        if ($board[$col + $row * 9] -eq $num) {return $false}
    }

    return $true
}

function Check-Region {
    param (
        [int[]]$board,
        [int]$pos,
        [int]$num
    )

    $row = [Math]::Truncate($pos / 9)
    $col = $pos % 9

    if ($row -le 2) {$row = 0}
    if (($row -ge 3) -and ($row -le 5)) {$row = 3}
    if ($row -ge 6) {$row = 6}

    if ($col -le 2) {$col = 0}
    if (($col -ge 3) -and ($col -le 5)) {$col = 3}
    if ($col -ge 6) {$col = 6}

    for ($i = 0; $i -lt 3; $i++) {
        for ($j = 0; $j -lt 3; $j++) {
            if ($board[($col + $j) + ($row + $i) * 9] -eq $num) {return $false}
        }
    }

    return $true
}

function Solve-Sudoku {
    param (
        [Parameter(Mandatory=$false)]
        [int[]]$board,
        [int]$pos = 0
    )

    while ($board[$pos] -ne 0) {
        $pos++
        if ($pos -ge $board.Count) {
            $global:success = $true
            return
        }
    }

    for ($i = 1; $i -lt 10; $i++) {
        if ((Check-Row $board $pos $i) -and (Check-Column $board $pos $i) -and (Check-Region $board $pos $i)) {
            $board[$pos] = $i
            Solve-Sudoku $board $pos
            if ($global:success) {return}
            $board[$pos] = 0
        }
    }
}

$boards = ((Get-Content -Path ".\sudoku.txt" -Raw) -split "Grid [0-9]+\s") | Select-Object -Skip 1 | ForEach-Object {($_ -replace "\s", "")} | ForEach-Object {,[int[]][string[]]$_.ToCharArray()}

for ($i = 0; $i -lt $boards.Count; $i++) {
    if ($i -le 9) {
        Write-Host "Grid $('{0:00}' -f $($i + 1))"
    } else {
        Write-Host "Grid $($i + 1)"
    }

    Solve-Sudoku $boards[$i]
    Print-Sudoku $boards[$i]

    $global:success = $false

    Write-Host
}

pause