function Add-Option {
    $options = @()

    for ($i = 0; $i -lt $args.Count; $i++) {
        $options += [PSCustomObject]@{Option=$args[$i]}
    }

    return $options
}

$smtp = Read-Host "Enter the SMTP server name (or IP)"
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

$users = Get-ADUser -Filter * -Properties Mobile, TelephoneNumber, EmailAddress, WhenCreated, Enabled, LockedOut | Select-Object Name, SamAccountName, Mobile, TelephoneNumber, EmailAddress, WhenCreated, Enabled, LockedOut | Sort-Object -Property @{e="LockedOut"; Descending=$true},@{e="Enabled"; Descending=$true},@{e="Name"; Descending=$false}

while ($user = $users | Out-GridView -Title "Users" -OutputMode Single) {
    $mailbox = Get-Mailbox -Identity $user.SamAccountName 2>$null
    if ($user.Enabled) {
        $option = "Disable User"
    } else {
        $option = "Enable User"
    }

    if ($mailbox) {
        $options = Add-Option "Unlock User" $option "Mailbox Info" "Mailbox Permissions" "Export Mailbox"
    } else {
        $options = Add-Option "Unlock User" $option
    }

    while ($choice = $options | Out-GridView -Title ($user.SamAccountName + " (" + $user.Name +")") -OutputMode Single) {
        switch ($choice.Option) {
            $options[0].Option {
                Unlock-ADAccount -Identity $user.SamAccountName
                break
            }

            $options[1].Option {
                if ($user.Enabled) {
                    Disable-ADAccount -Identity $user.SamAccountName
                    $options[1].Option = "Enable User"
                    $user.Enabled = $false
                } else {
                    Enable-ADAccount -Identity $user.SamAccountName
                    $options[1].Option = "Disable User"
                    $user.Enabled = $true
                }

                break
            }

            $options[2].Option {
                $mailboxStats = $mailbox | Get-MailboxStatistics
                $folderStats = $mailbox | Get-MailboxFolderStatistics

                $foldersCount = $folderStats | Measure-Object | Select-Object Count

                $sendConnectorLimits = Get-SendConnector | Where-Object {$_.Enabled -eq $true}
                $receiveConnectorLimits = Get-ReceiveConnector | Where-Object {($_.TransportRole -eq "HubTransport") -and ($_.Identity -like "*Default*")} | Select-Object -First 1
                $transportLimits = Get-TransportConfig

                $recepientLimit = @($mailbox.RecipientLimits, $transportLimits.MaxRecipientEnvelopeLimit) | Where-Object {$_ -ne "Unlimited"} | Select-Object -First 1
                if (!$recepientLimit) {$recepientLimit = $receiveConnectorLimits.MaxRecipientsPerMessage}
                $maxSendSize = @($mailbox.MaxSendSize, $transportLimits.MaxSendSize, $sendConnectorLimits.MaxMessageSize) | Where-Object {$_ -ne "Unlimited"} | Sort-Object | Select-Object -First 1
                $maxReceiveSize = @($mailbox.MaxReceiveSize, $transportLimits.MaxReceiveSize, $receiveConnectorLimits.MaxMessageSize) | Where-Object {$_ -ne "Unlimited"} | Sort-Object | Select-Object -First 1

                $limitWarning = @($mailboxStats.DatabaseIssueWarningQuota.Value, $mailbox.IssueWarningQuota) | Sort-Object | Select-Object -First 1
                $sendLimit = @($mailboxStats.DatabaseProhibitSendQuota.Value, $mailbox.ProhibitSendQuota) | Sort-Object | Select-Object -First 1
                $sendReceiveLimit = @($mailboxStats.DatabaseProhibitSendReceiveQuota.Value, $mailbox.ProhibitSendReceiveQuota) | Sort-Object | Select -First 1

                $stats = [ordered]@{
                    "Email Address" = $mailbox.WindowsEmailAddress
                    "Database" = $mailbox.Database
                    "Recipient Limit" = $recepientLimit
                    "Folder Count" = $foldersCount.Count
                    "Item Count" = $mailboxStats.ItemCount
                    "Used Space" = $mailboxStats.TotalItemSize
                    "Limit Warning" = $limitWarning
                    "Send Limit" = $sendLimit
                    "Send & Receive Limit" = $sendReceiveLimit
                    "Max Send Size" = $maxSendSize
                    "Max Receive Size" = $maxReceiveSize
                }

                $result = $stats | Out-GridView -Title ($user.SamAccountName + " (" + $user.Name +")") -PassThru
                $filePath = $user.Name.Replace("""", "")
                if ($result) {$result | Export-Csv -Path ".\$($filePath).csv" -Encoding UTF8 -NoTypeInformation}

                break
            }

            $options[3].Option {
                $calendar = Get-MailboxFolderStatistics -Identity $user.SamAccountName -FolderScope Calendar | Where-Object {$_.FolderType -eq "Calendar"} | Select-Object -ExpandProperty Name
                $tasks = Get-MailboxFolderStatistics -Identity $user.SamAccountName -FolderScope Tasks | Where-Object {$_.FolderType -eq "Tasks"} | Select-Object -ExpandProperty Name
                $inbox = Get-MailboxFolderStatistics -Identity $user.SamAccountName -FolderScope Inbox | Where-Object {$_.FolderType -eq "Inbox"} | Select-Object -ExpandProperty Name
                $contacts = Get-MailboxFolderStatistics -Identity $user.SamAccountName -FolderScope Contacts | Where-Object {$_.FolderType -eq "Contacts"} | Select-Object -ExpandProperty Name
                $notes = Get-MailboxFolderStatistics -Identity $user.SamAccountName -FolderScope Notes | Where-Object {$_.FolderType -eq "Notes"} | Select-Object -ExpandProperty Name

                $delegates = @()

                $mailboxes = Get-MailboxPermission -Identity $user.SamAccountName | Where-Object {($_.User -ne "NT AUTHORITY\SELF") -and ($_.IsInherited -eq $false)} | Select-Object User, AccessRights
                foreach ($box in $mailboxes) {
                    $samName = $box.User
                    if ($samName.Contains("\")) {$samName = $box.User.Split("\")[1]}
                    $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
                    $delegates += [PSCustomObject]@{User=($samName + " " + $displayName); FolderName="*"; Rights=$box.AccessRights[0]}
                }

                $mailboxes = $mailbox | Get-ADPermission | Where-Object {($_.User -ne "NT AUTHORITY\SELF") -and ($_.IsInherited -eq $false) -and ($_.ExtendedRights -eq "Send-As")}
                foreach ($box in $mailboxes) {
                    $samName = $box.User
                    if ($samName.Contains("\")) {$samName = $box.User.Split("\")[1]}
                    $displayName = "(" + (&{if ($name = (Get-ADObject -Filter {SamAccountName -eq $samName}).Name) {$name} else {"?"}}) + ")"
                    $delegates += [PSCustomObject]@{User=($samName + " " + $displayName); FolderName="-"; Rights="Send As"}
                }

                $mailboxes = $mailbox | ForEach-Object {$_.GrantSendOnBehalfTo | ForEach-Object {$_.Split("/")[4]}}
                foreach ($box in $mailboxes) {
                    $samName = &{if ($name = (Get-ADObject -Filter "Name -eq ""$($box)""" -Properties SamAccountName).SamAccountName) {$name} else {"?"}}
                    $displayName = "(" + $box + ")"
                    $delegates += [PSCustomObject]@{User=($samName + " " + $displayName); FolderName="-"; Rights="Send on Behalf"}
                }
                

                $folders = @($calendar, $tasks, $inbox, $contacts, $notes)
                foreach ($folder in $folders) {
                    $permissions = Get-MailboxFolderPermission -Identity "$($user.SamAccountName):\$folder" | Where-Object {($_.User.DisplayName -ne "Default") -and ($_.AccessRights -ne "None")} | Select-Object User, FolderName, AccessRights
                    foreach ($permission in $permissions) {
                        $samName = &{if ($name = (Get-ADObject -Filter "DisplayName -eq ""$($permission.User)""" -Properties SamAccountName).SamAccountName) {$name} else {"?"}}
                        $displayName = "(" + $permission.User + ")"
                        $delegates += [PSCustomObject]@{User=($samName + " " + $displayName); FolderName=$permission.FolderName; Rights=$permission.AccessRights[0]}
                    }
                }

                $result = $delegates | Out-GridView -Title ($user.SamAccountName + " (" + $user.Name +")") -PassThru
                $filePath = $user.Name.Replace("""", "")
                if ($result) {$result | Export-Csv -Path ".\$($filePath).csv" -Encoding UTF8 -NoTypeInformation}

                break
            }

            $options[4].Option {
                $options = Add-Option "All" "Inbox" "Semt Items" "Calendar"

                while ($choice = $options | Out-GridView -Title ($user.SamAccountName + " (" + $user.Name +")") -OutputMode Single) {
                    $path = Read-Host "Enter the path for the PST file"

                    switch ($choice.Option) {
                        $options[0].Option {
                            $filePath = $user.Name.Replace("""", "")
                            $request = New-MailboxExportRequest -Mailbox $user.SamAccountName -FilePath $path
                            break
                        }

                        $options[1].Option {
                            $filePath = $user.Name.Replace("""", "")
                            $request = New-MailboxExportRequest -Mailbox $user.SamAccountName -IncludeFolders "#Inbox#" -FilePath $path
                            break
                        }

                        $options[2].Option {
                            $filePath = $user.Name.Replace("""", "")
                            $request = New-MailboxExportRequest -Mailbox $user.SamAccountName -IncludeFolders "#SentItems#" -FilePath $path
                            break
                        }

                        $options[3].Option {
                            $filePath = $user.Name.Replace("""", "")
                            $request = New-MailboxExportRequest -Mailbox $user.SamAccountName -IncludeFolders "#Calendar#" -FilePath $path
                            break
                        }
                    }

                    if ($request) {
                        $requestStats = Get-MailboxExportRequestStatistics -Identity $request.RequestGuid

                        while ($request.Status -ne "Completed") {
                            Write-Progress -Activity "Exporting PST..." -Status "$($requestStats.PercentComplete)% out of 100%." -PercentComplete $requestStats.PercentComplete
                            Start-Sleep -Seconds 5

                            $request = Get-MailboxExportRequest -Identity $request.RequestGuid
                            $requestStats = Get-MailboxExportRequestStatistics -Identity $request.RequestGuid
                        }
                        Write-Progress -Activity "Completed" -Completed

                        Get-MailboxExportRequest -Identity $request.RequestGuid | Remove-MailboxExportRequest -Confirm:$false
                        $request = ""
                    }
                }

                $options = Add-Option "Unlock User" $option "Mailbox Info" "Export Mailbox"
                break
            }
        }
    }
}