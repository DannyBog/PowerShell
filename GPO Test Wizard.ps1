$title = "GPO Test Wizard"
$all = New-Object System.Management.Automation.Host.ChoiceDescription "&All", "Create a test group with users and computers."
$computers = New-Object System.Management.Automation.Host.ChoiceDescription "&Computers", "Create a test group with computers only."
$users = New-Object System.Management.Automation.Host.ChoiceDescription "&Users", "Create a test group users only."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($all, $computers, $users)
$response = $host.UI.PromptForChoice($title, $null, $options, 0)

$dn = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName
$subnets = ((Read-Host "Enter the subnets (comma seperated) that will participate in the test") -replace '\s+', '').Split(",")
$testGroup = Get-ADObject -Filter {Name -eq "GPO Test" -and ObjectClass -eq "Group"}
if (!$testGroup) {$testGroup = New-ADGroup -Name "GPO Test" -GroupCategory Security -GroupScope Global -DisplayName "GPO Test" -Path "OU=Groups,$dn" -Description "Members of this group are being tested by a new GPO" -PassThru}

$computers = Get-ADComputer -Filter * -SearchBase "OU=Computers,$dn" -Properties Description

Write-Host

foreach ($computer in $computers) {
    try {
        $computerObj = Get-ADComputer -Identity $computer.Name -Properties Description
    } catch {
        $computerObj = ""
    }

    if ($computerObj) {
        $description = $computerObj.Description

        if ($description) {
            $user = $description.Split("()")[1]
            $ips = $description.Split("()")[5].SubString(1) -replace "\s+", ","
            $ips = $ips.SubString(0, $ips.Length - 1).Split(",") | Where-Object {$_ -like "1.1.*"}
            foreach ($ip in $ips) {
                foreach ($subnet in $subnets) {
                    if ($ip.Split(".")[2] -eq $subnet) {
                        switch ($response) {
                            0 {
                                Add-ADGroupMember -Identity $testGroup -Members $computer, $user
                                break
                            }

                            1 {
                                Add-ADGroupMember -Identity $testGroup -Members $computer
                                break
                            }

                            2 {
                                Add-ADGroupMember -Identity $testGroup -Members $user
                                break
                            }
                        }
                    }
                }
            }
        }
    }
}