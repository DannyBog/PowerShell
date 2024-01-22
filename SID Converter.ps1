$nameOrSID = Read-Host "Enter an Active Directory object name (or SID)"

if ($nameOrSID -like "S-1-5-21-*") {
    try {$objSID = New-Object System.Security.Principal.SecurityIdentifier($nameOrSID)} catch {}

    if ($objSID) {
        try {$objUser = $objSID.Translate([System.Security.Principal.NTAccount])} catch {Write-Host "User not found :("}
        if ($objUser) {Write-Host $objUser.Value.TrimEnd("$")}
    } else {
        Write-Host "User not found :("
    }
} else {
    $obj = New-Object System.Security.Principal.NTAccount($env:USERDOMAIN, $nameOrSID)
    try {$objSID = $obj.Translate([System.Security.Principal.SecurityIdentifier])} catch {}
    if (!$objSID) {
        $obj = New-Object System.Security.Principal.NTAccount($env:USERDOMAIN, "$nameOrSID`$")
        try {$objADComputerSID = $obj.Translate([System.Security.Principal.SecurityIdentifier])} catch {}

        if ($objADComputerSID) {
            Write-Host $objADComputerSID.Value "(AD Computer Object SID)"

            $objLocalComputer = (Get-WmiObject -ComputerName $nameOrSID -Query "SELECT SID FROM Win32_UserAccount WHERE LocalAccount = 'True'" | Select-Object -First 1 -ExpandProperty SID).Split("-")
            if ($objLocalComputer) {
                $objLocalComputerSID = $objLocalComputer[0..($objLocalComputer.Length - 2)] -join "-"
                Write-Host $objLocalComputerSID "(Local Computer SID)" 
            }
        } else {
            Write-Host "SID not found :("            
        }
    } else {
        Write-Host $objSID.Value
    }
}

pause