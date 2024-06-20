$groupCustomObj = Get-ADGroup -Filter * | ForEach-Object {[PSCustomObject]@{Name=$_ | Select-Object -ExpandProperty Name; SamAccountName=$_ | Select-Object SamAccountName}}
$groupCustomObj = $groupCustomObj | Sort-Object Name

while ($choice = $groupCustomObj | Select-Object Name | Out-GridView -Title "Groups" -OutputMode Single) {
    $groupName = $choice.Name
    $group = (($groupCustomObj | Where-Object {($_.Name -eq $groupName)}).SamAccountName).psobject.Properties.value
    $members = Get-ADGroup -Identity $group -Properties Members | Select-Object -ExpandProperty Members

    [System.Collections.ArrayList]$entries = @()
    foreach ($member in $members) {
        $object = (Get-Culture).TextInfo.ToTitleCase((Get-ADObject -Filter {DistinguishedName -eq $member}).ObjectClass)

        switch ($object) {
            "User" {
                $member = $member | Get-ADUser -Properties Mobile, EmailAddress | Select-Object Name, SamAccountName, Mobile, EmailAddress
                $memberCustomObj = [PSCustomObject]@{Type=$object; Name=$member.Name; Username=$member.SamAccountName; Mobile=$member.Mobile; EmailAddress=$member.EmailAddress}
                $entries.Add($memberCustomObj) | Out-Null
                break
            }

            "Group" {
                $member = $member | Get-ADGroup | Select-Object -ExpandProperty Name
                $memberCustomObj = [PSCustomObject]@{Type=$object; Name=$member; Username=""; Mobile=""; EmailAddress=""}
                $entries.Add($memberCustomObj) | Out-Null
                break
            }

            "Computer" {
                $member = $member | Get-ADComputer | Select-Object -ExpandProperty Name 
                $memberCustomObj = [PSCustomObject]@{Type=$object; Name=$member; Username=""; Mobile=""; EmailAddress=""}
                $entries.Add($memberCustomObj) | Out-Null
                break
            }

            "Contact" {
                $member = $member | Get-ADObject -Properties Mobile, Mail | Select-Object Name, Mobile, Mail
                $memberCustomObj = [PSCustomObject]@{Type=$object; Name=$member.Name; Username=""; Mobile=$member.Mobile; EmailAddress=$member.Mail}
                $entries.Add($memberCustomObj) | Out-Null
                break
            }

            default {
                $memberCustomObj = [PSCustomObject]@{Type="?"; Name=$member.Split("=,")[1]; Username="?"; Mobile="?"; EmailAddress="?"}
                $entries.Add($memberCustomObj) | Out-Null
            }
        }
    }

    $result = $entries | Sort-Object -Property @{e="Type"; Descending=$true}, @{e="Name"; Descending=$false} | Out-GridView -Title $groupName -PassThru
    if ($result) {$result | Export-Csv -Path ".\$groupName.csv" -Encoding UTF8 -NoTypeInformation}
}