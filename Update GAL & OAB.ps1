$smtp = Read-Host "Enter the SMTP server name (or IP)"

$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

Get-GlobalAddressList -DomainController $env:LOGONSERVER.Substring(2) | Update-GlobalAddressList -DomainController $env:LOGONSERVER.Substring(2)
Get-OfflineAddressBook -DomainController $env:LOGONSERVER.Substring(2) | Update-OfflineAddressBook -DomainController $env:LOGONSERVER.Substring(2)