$smtp = Read-Host "Enter the SMTP server name (or IP)"

do {
    $mailbox = Read-Host "Enter an email address"
} while ($mailbox -eq "")

$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

$deletedMails = Search-MailboxAuditLog -Identity $mailbox -LogonTypes Admin, Delegate, Owner -ShowDetails | Where-Object {$_.Operation -eq "MoveToDeletedItems" -or $_.Operation -eq "SoftDelete" -or $_.Operation -eq "HardDelete"} | Select-Object @{n="Operation"; e={$_.Operation}}, @{n="Type"; e={$_.LogonType}}, @{n="User"; e={$_.LogonUserDisplayName}}, @{n="IP"; e={$_.ClientIPAddress}}, @{n="Subject"; e={$_.SourceItemSubjectsList}}, @{n="Source Folder"; e={$_.SourceItemfolderPathNamesList}}, @{n="Date"; e={$_.LastAccessed}}
if ($deletedMails) {$deletedMails | Export-Csv -Path ".\Result.csv" -Encoding UTF8 -NoTypeInformation}