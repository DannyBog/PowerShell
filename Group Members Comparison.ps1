$groupCustomObj = Get-ADGroup -Filter * | ForEach-Object {[PSCustomObject]@{Name=$_ | Select-Object -ExpandProperty Name; SamAccountName=$_ | Select-Object SamAccountName}}

$choice1 = Get-ADGroup -Filter * | Select-Object Name | Sort-Object Name | Out-GridView -Title "Groups 1" -OutputMode Single
$group1 = (($groupCustomObj | Where-Object {($_.Name -eq $choice1.Name)}).SamAccountName).psobject.Properties.value

$choice2 = Get-ADGroup -Filter * | Select-Object Name | Sort-Object Name | Out-GridView -Title "Groups 2" -OutputMode Single
$group2 = (($groupCustomObj | Where-Object {($_.Name -eq $choice2.Name)}).SamAccountName).psobject.Properties.value

$membersGroup1 = Get-ADGroupMember -Identity $group1
$membersGroup2 = Get-ADGroupMember -Identity $group2
$members = Compare-Object -ReferenceObject $membersGroup1 -DifferenceObject $membersGroup2 -IncludeEqual

$entries = @()

foreach ($member in $members) {
    switch ($member.SideIndicator) {
        "<=" {
            $entries += [PSCustomObject]@{$group1=$member.InputObject.name; $group2=""}
            break
        }

        "=>" {
            $entries += [PSCustomObject]@{$group1=""; $group2=$member.InputObject.name}
            break
        }

        "==" {
            $entries += [PSCustomObject]@{$group1=$member.InputObject.name; $group2=$member.InputObject.name}
            break
        }
    }    
}

if ($entries) {$entries | Export-Csv -Path ".\Comparison.csv" -Encoding UTF8 -NoTypeInformation}