function Add-Option {
    $options = @()

    for ($i = 0; $i -lt $args.Count; $i++) {
        $options += [PSCustomObject]@{Option=$args[$i]}
    }

    return $options
}

$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

$users = Get-ADUser -Filter {msExchMailboxGuid -like "*"} -Properties Mail | Select-Object Name, SamAccountName, Mail | Sort-Object Name

while ($user1 = $users | Out-GridView -Title "Mailbox to search" -OutputMode Single) {
    $options = Add-Option "Search for messages sent to" "Search for messages received from"

    while ($choice = $options | Out-GridView -Title ($user1.SamAccountName + " (" + $user1.Name +")") -OutputMode Single) {
        switch ($choice.Option) {
            $options[0].Option {
                while ($user2 = $users | Out-GridView -Title "Search for messages sent to" -OutputMode Single) {
                    Write-Progress -Activity "Gathering Data..."
                    $logs = Search-MessageTrackingReport -Identity $user1.Mail -Recipients $user2.Mail -BypassDelegateChecking
                    
                    $entries = @()
                    foreach ($log in $logs) {
                        $entries += [PSCustomObject]@{From=$log.FromDisplayName; To=$log.RecipientDisplayNames -join ", "; Subject=$log.Subject; "Sent Time"=$log.SubmittedDateTime.AddHours(2)}
                    }
                    Write-Progress -Activity "Completed" -Completed

                    $results = $entries | Out-GridView -Title ("Messages sent from: " + $user1.Name + " to: " + $user2.Name) -PassThru
                    if ($results) {$results | Export-Csv -Path ".\Delivery Report.csv" -Encoding UTF8 -NoTypeInformation}
                }

                break
            }

            $options[1].Option {
                while ($user2 = $users | Out-GridView -Title "Search for messages received from" -OutputMode Single) {
                    Write-Progress -Activity "Gathering Data..."
                    $logs = Search-MessageTrackingReport -Identity $user1.Mail -Sender $user2.Mail -BypassDelegateChecking
                    
                    $entries = @()
                    foreach ($log in $logs) {
                        $entries += [PSCustomObject]@{From=$log.FromDisplayName; To=$log.RecipientDisplayNames -join ", "; Subject=$log.Subject; "Sent Time"=$log.SubmittedDateTime.AddHours(2)}
                    }
                    Write-Progress -Activity "Completed" -Completed

                    $results = $entries | Out-GridView -Title ("Messages sent to: " + $user1.Name + " from: " + $user2.Name) -PassThru
                    if ($results) {$results | Export-Csv -Path ".\Delivery Report.csv" -Encoding UTF8 -NoTypeInformation}
                }

                break
            }
        }
    }
}