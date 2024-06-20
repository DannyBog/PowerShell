$objects = Get-ADObject -SearchBase (Get-ADDomain).DeletedObjectsContainer -IncludeDeletedObjects -Filter * -Properties *

$entries = @()

foreach ($object in $objects) {
    switch ($object.ObjectClass) {
        "User" {
            $entries += [PSCustomObject]@{Class="User"; Name=$object."msDS-LastKnownRDN"; IP=""; Username=$object.SamAccountName; Mobile=$object.Mobile; "When Deleted"=$object.WhenChanged}
            break
        }

        "Group" {
            $entries += [PSCustomObject]@{Class="Group"; Name=$object."msDS-LastKnownRDN"; IP=""; Username=""; Mobile=""; "When Deleted"=$object.WhenChanged}
            break
        }

        "GroupPolicyContainer" {
            $entries += [PSCustomObject]@{Class="Group Policy"; Name=$object.DisplayName; IP=""; Username=""; Mobile=""; "When Deleted"=$object.WhenChanged}
            break
        }

        "Computer" {
            if ($object.Description) {
                $ip = $object.Description.Split("()")[5].SubString(1) -replace "\s+", ", "
                $ip = $ip.SubString(0, $ip.Length - 2)
            }

            $entries += [PSCustomObject]@{Class="Computer"; Name=$object."msDS-LastKnownRDN"; IP=$ip; Username=""; Mobile=""; "When Deleted"=$object.WhenChanged}
            break
        }

        "Contact" {
            $entries += [PSCustomObject]@{Class="Contact"; Name=$object."msDS-LastKnownRDN"; IP=""; Username=""; Mobile=$object.Mobile; "When Deleted"=$object.WhenChanged}
            break
        }

        "Container" {
            $entries += [PSCustomObject]@{Class="Container"; Name=$object."msDS-LastKnownRDN"; IP=""; Username=""; Mobile=$object.Mobile; "When Deleted"=$object.WhenChanged}
            break
        }

        "PrintQueue" {
            $ip = $object | Select-Object -ExpandProperty PortName
            $ip = $ip.Split("_")[0]

            $entries += [PSCustomObject]@{Class="Printer"; Name=$object."msDS-LastKnownRDN"; IP=$ip; Username=""; Mobile=""; "When Deleted"=$object.WhenChanged}
            break
        }

        default {
            $entries += [PSCustomObject]@{Class=$object.ObjectClass; Name=$object."msDS-LastKnownRDN"; IP=""; Username=""; Mobile=""; "When Deleted"=$object.WhenChanged}
        }
    }

    $ip = ""
}

while ($choice = $entries | Sort-Object Class, Name | Out-GridView -Title "Recycle Bin" -OutputMode Single) {
    $choice = Get-ADObject -SearchBase (Get-ADDomain).DeletedObjectsContainer -IncludeDeletedObjects -Filter "msDS-LastKnownRDN -eq `"$($choice.Name)`""
    Restore-ADObject -Identity $choice
}