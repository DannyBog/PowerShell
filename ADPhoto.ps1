do {
    $username = Read-Host "Enter a username"
    $existingUsername = Get-ADUser -Filter {SamAccountName -eq $username}
    if (!$existingUsername) {Write-Error "Username does not exist."}
} while (!$existingUsername)

Get-ADUser -Identity $username -Properties * | Select-Object -ExpandProperty thumbnailPhoto | Add-Content -Path ".\$($username).jpg" -Encoding Byte