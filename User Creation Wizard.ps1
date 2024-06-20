$title = "User Creation Wizard"
$copyU = New-Object System.Management.Automation.Host.ChoiceDescription "&Copy User", "Copy an existing user's attributes over to a new user."
$singleU = New-Object System.Management.Automation.Host.ChoiceDescription "&Single User", "Create a single user."
$singleUM = New-Object System.Management.Automation.Host.ChoiceDescription "S&ingle Exchange User", "Create a single user with a mailbox."
$multipleU = New-Object System.Management.Automation.Host.ChoiceDescription "&Multiple Users", "Create multiple users."
$multipleUM = New-Object System.Management.Automation.Host.ChoiceDescription "M&ultiple Exchange Users", "Create multiple users with mailboxes."
$file = New-Object System.Management.Automation.Host.ChoiceDescription "&File", "Import users from a CSV file."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($copyU, $singleU, $singleUM, $multipleU, $multipleUM, $file)
$response = $host.UI.PromptForChoice($title, $null, $options, 0)
$dn = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName

Write-Host

switch ($response) {
    0 {
        $smtp = Read-Host "Enter the SMTP server name (or IP)"
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
        Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

        $username = Read-Host "Enter a username"
        $password = Read-Host "Enter a password"
        $firstName = Read-Host "Enter a firstname"
        $lastName = Read-Host "Enter a lastname"
        $id = Read-Host "Enter an ID"
        $mobileNum = Read-Host "Enter a mobile number"

        if ($firstName -and $lastName) {
            $displayName = $firstName + " " + $lastName
        } elseif ($firstName) {
            $displayName = $firstName
        } else {
            $displayName = $username
        }

        if ($id) {
            $upn = "$id@$env:USERDNSDOMAIN"
        } else {
            $upn = "$username@$env:USERDNSDOMAIN"
        }

        do {
            $existingUsername = Read-Host "Enter an existing username"
            $existingUsername = Get-ADUser -Filter {SamAccountName -eq $existingUsername}
            if (!$existingUsername) {Write-Error "Username does not exist."}
        } while (!$existingUsername)

        $existingMailbox = Get-Mailbox -Identity $existingUsername.Name 2>$null
        if ($existingMailbox) {
            do {
                $email = Read-Host "Enter an email address"
                $existingEmail = Get-ADUser -Filter {EmailAddress -eq $email} -Properties EmailAddress
                if ($existingEmail) {Write-Error "Email already exists ($($existingEmail.Name) [$($existingEmail.EmailAddress)])"}
            } while ($existingEmail)
            $alias = $email.Split("@")[0]

            $user = New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$displayName" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$displayName" -SamAccountName $username -UserPrincipalName $upn -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -ChangePasswordAtLogon $true -Enabled $true -PassThru
            Enable-Mailbox -DomainController $env:LOGONSERVER.Substring(2) -Identity $user.ObjectGUID.Guid -Alias $alias -Database $existingMailbox.Database
        } else {
            $email = Read-Host "Enter an external email address"
            New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$displayName" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$displayName" -SamAccountName $username -UserPrincipalName $upn -EmailAddress $email -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Enabled $true
        }

        if ($mobileNum) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -MobilePhone $mobileNum}
        if ($id) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -EmployeeID $id}

        $properties = Get-ADUser -Identity $existingUsername -Properties * | Select-Object Description, Office, StreetAddress, l, ScriptPath, HomeDrive, HomeDirectory, Department, Company, Manager

        if ($properties.Description) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -Description $properties.Description}
        if ($properties.Office) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -Office $properties.Office}
        if ($properties.StreetAddress) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -StreetAddress $properties.StreetAddress}
        if ($properties.l) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -City $properties.l}
        if ($properties.ScriptPath) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -ScriptPath $properties.ScriptPath}
        if ($properties.HomeDrive) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -HomeDrive $properties.HomeDrive}
        if ($properties.HomeDirectory) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -HomeDirectory ((Split-Path -Path $properties.HomeDirectory -Parent) + "\$username")}
        if ($properties.Department) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -Department $properties.Department}
        if ($properties.Company) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -Company $properties.Company}
        if ($properties.Manager) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -Manager $properties.Manager}

        Get-ADUser -Identity $existingUsername -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Add-ADGroupMember -Server $env:LOGONSERVER.Substring(2) -Members $username

        break
    }

    1 {
        $username = Read-Host "Enter a username"
        $password = Read-Host "Enter a password"
        $firstName = Read-Host "Enter a firstname"
        $lastName = Read-Host "Enter a lastname"
        $id = Read-Host "Enter an ID"
        $mobileNum = Read-Host "Enter a mobile number"

        if ($firstName -and $lastName) {
            $displayName = $firstName + " " + $lastName
        } elseif ($firstName) {
            $displayName = $firstName
        } else {
            $displayName = $username
        }

        if ($id) {
            $upn = "$id@$env:USERDNSDOMAIN"
        } else {
            $upn = "$username@$env:USERDNSDOMAIN"
        }

        $email = Read-Host "Enter an external email address"

        New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$displayName" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$displayName" -SamAccountName "$username" -UserPrincipalName "$upn" -EmailAddress $email -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Enabled $true
        if ($mobileNum) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -MobilePhone $mobileNum}
        if ($id) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -EmployeeID $id}

        break
    }

    2 {
        $smtp = Read-Host "Enter the SMTP server name (or IP)"
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
        Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

        $username = Read-Host "Enter a username"
        $password = Read-Host "Enter a password"
        $firstName = Read-Host "Enter a firstname"
        $lastName = Read-Host "Enter a lastname"
        $id = Read-Host "Enter an ID"
        $mobileNum = Read-Host "Enter a mobile number"

        if ($firstName -and $lastName) {
            $displayName = $firstName + " " + $lastName
        } elseif ($firstName) {
            $displayName = $firstName
        } else {
            $displayName = $username
        }

        if ($id) {
            $upn = "$id@$env:USERDNSDOMAIN"
        } else {
            $upn = "$username@$env:USERDNSDOMAIN"
        }

        do {
            $email = Read-Host "Enter an email address"
            $existingEmail = Get-ADUser -Filter {EmailAddress -eq $email} -Properties EmailAddress
            if ($existingEmail) {Write-Host "Email already exists ($($existingEmail.Name) [$($existingEmail.EmailAddress)])"}
        } while ($existingEmail)
        $alias = $email.Split("@")[0]

        $database = Get-MailboxDatabase | Select-Object Name | Out-GridView -Title "Databases" -OutputMode Single

        $user = New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$displayName" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$displayName" -SamAccountName $username -UserPrincipalName $upn -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -ChangePasswordAtLogon $true -Enabled $true -PassThru
        Enable-Mailbox -DomainController $env:LOGONSERVER.Substring(2) -Identity $user.ObjectGUID.Guid -Alias $alias -Database $database.Name

        if ($mobileNum) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -MobilePhone $mobileNum}
        if ($id) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username -EmployeeID $id}

        break
    }

    3 {
        $username = Read-Host "Enter a username"
        $password = Read-Host "Enter a password"
        $usersNum = Read-Host "Enter the number of users to create"

        $username += "*"
        $start = ((Get-ADUser -Filter {SamAccountName -like $username} | Select-Object Name) -replace "[^0-9]" | Measure-Object -Maximum).Maximum + 1
        $end = $start + $usersNum - 1
        $username = $username.TrimEnd("*")

        for ($i = $start; $i -le $end; $i++) {
            New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$username$i" -GivenName "$username$i" -Surname "$username$i" -DisplayName "$username$i" -SamAccountName "$username$i" -UserPrincipalName "$username$i@$env:USERDNSDOMAIN" -ScriptPath "logon.bat" -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Enabled $true
        }

        break
    }

    4 {
        $smtp = Read-Host "Enter the SMTP server name (or IP)"
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
        Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

        $username = Read-Host "Enter a username"
        $password = Read-Host "Enter a password"
        $usersNum = Read-Host "Enter the number of users to create"

        $database = Get-MailboxDatabase | Select-Object Name | Out-GridView -Title "Databases" -OutputMode Single

        $username += "*"
        $start = ((Get-ADUser -Filter {SamAccountName -like $username} | Select-Object SamAccountName) -replace "[^0-9]" | Measure-Object -Maximum).Maximum + 1
        $end = $start + $usersNum - 1
        $username = $username.TrimEnd("*")

        for ($i = $start; $i -le $end; $i++) {
            $user = New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$username$i" -GivenName "$username$i" -Surname "$username$i" -DisplayName "$username$i" -SamAccountName "$username$i" -UserPrincipalName "$username$i@$env:USERDNSDOMAIN" -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Enabled $true -PassThru
            Enable-Mailbox -DomainController $env:LOGONSERVER.Substring(2) -Identity $user.ObjectGUID.Guid -Alias $username$i -Database $database.Name

            Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $username$i -ScriptPath "logon.bat"
        }

        break
    }

    5 {
        $smtp = Read-Host "Enter the SMTP server name (or IP)"
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$smtp/PowerShell/" -Authentication Default
        Import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null

        $users = Import-Csv -Path ".\users.csv"

        foreach ($user in $users) {
            if ($user.FirstName -and $user.LastName) {
                $displayName = $user.FirstName + " " + $user.LastName
                $adUser = New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$displayName" -GivenName "$($user.FirstName)" -Surname "$($user.LastName)" -DisplayName "$displayName" -SamAccountName $user.Username -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $user.Password -AsPlainText -Force) -Enabled $true -PassThru
            } elseif ($user.FirstName) {
                $adUser = New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name "$($user.FirstName)" -GivenName "$($user.FirstName)" -DisplayName "$displayName" -SamAccountName $user.Username -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $user.Password -AsPlainText -Force) -Enabled $true -PassThru
            } else {
                $adUser = New-ADUser -Server $env:LOGONSERVER.Substring(2) -Name $user.Username -DisplayName $user.Username -SamAccountName $user.Username -Path "OU=Users,$dn" -AccountPassword (ConvertTo-SecureString $user.Password -AsPlainText -Force) -Enabled $true -PassThru
            }

            if ($user.ID) {
                $upn = "$($user.ID)@$env:USERDNSDOMAIN"
                Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -UserPrincipalName $upn -EmployeeID $user.ID
            } else {
                $upn = "$($user.Username)@$env:USERDNSDOMAIN"
                Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -UserPrincipalName $upn
            }

            if ($user.Description) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -Description $user.Description}
            if ($user.Office) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -Description $user.Office}
            if ($user.StreetAddress) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -StreetAddress $user.StreetAddress}
            if ($user.City) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -City $user.City}
            if ($user.ScriptPath) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -ScriptPath $user.ScriptPath}
            if ($user.HomeDrive) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -HomeDrive $user.HomeDrive}
            if ($user.HomeDirectory) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -HomeDirectory ((Split-Path -Path $user.HomeDirectory -Parent) + "\$($user.Username)")}
            if ($user.Mobile) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -MobilePhone $user.Mobile}
            if ($user.Title) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -Title $user.Title}
            if ($user.Department) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -Department $user.Department}
            if ($user.Company) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -Company $user.Company}
            if ($user.Manager) {Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -Manager $user.Manager}

            if ($user.Email) {
                $domain = $user.Email.Split("@")[1]
                if ($domain -eq $env:USERDNSDOMAIN) {
                    $alias = $user.Email.Split("@")[0]
                    $database = Get-MailboxDatabase | Select-Object Name | Out-GridView -Title "Databases" -OutputMode Single
                    Enable-Mailbox -DomainController $env:LOGONSERVER.Substring(2) -Identity $adUser.ObjectGUID.Guid -Alias $alias -Database $database.Name
                } else {
                    Set-ADUser -Server $env:LOGONSERVER.Substring(2) -Identity $adUser -EmailAddress $user.Email
                }
            }
        }

        break
    }
}