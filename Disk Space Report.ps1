$servers = Get-Content ".\servers.txt" | Where-Object {$_.Trim() -ne ""}
$freeSpaceFileName = ".\servercheck.htm"

$warning = 25
$critical = 15

New-Item -ItemType File $freeSpaceFileName -Force | Out-Null

Function WriteHTMLHeader {
    param ($fileName)

    $date = (Get-Date).ToString("dd/MM/yyyy")

    Add-Content $fileName "<html>"
    Add-Content $fileName "<head>"
    Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
    Add-Content $fileName '<title>2012 Servers Disk Space Report</title>'
    Add-Content $fileName '<STYLE TYPE="text/css">'
    Add-Content $fileName  "<!--"
    Add-Content $fileName  "td {"
    Add-Content $fileName  "font-family: Tahoma;"
    Add-Content $fileName  "font-size: 11px;"
    Add-Content $fileName  "border-top: 1px solid #999999;"
    Add-Content $fileName  "border-right: 1px solid #999999;"
    Add-Content $fileName  "border-bottom: 1px solid #999999;"
    Add-Content $fileName  "border-left: 1px solid #999999;"
    Add-Content $fileName  "padding-top: 0px;"
    Add-Content $fileName  "padding-right: 0px;"
    Add-Content $fileName  "padding-bottom: 0px;"
    Add-Content $fileName  "padding-left: 0px;"
    Add-Content $fileName  "}"
    Add-Content $fileName  "body {"
    Add-Content $fileName  "margin-left: 5px;"
    Add-Content $fileName  "margin-top: 5px;"
    Add-Content $fileName  "margin-right: 0px;"
    Add-Content $fileName  "margin-bottom: 10px;"
    Add-Content $fileName  ""
    Add-Content $fileName  "table {"
    Add-Content $fileName  "border: thin solid #000000;"
    Add-Content $fileName  "}"
    Add-Content $fileName  "-->"
    Add-Content $fileName  "</style>"
    Add-Content $fileName  "</head>"
    Add-Content $fileName  "<body>"
    Add-Content $fileName  "<table width='100%'>"
    Add-Content $fileName  "<tr bgcolor='#CCCCCC'>"
    Add-Content $fileName  "<td colspan='7' height='25' align='center'>"
    Add-Content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Servers DiskSpace Report - $date</strong></font>"
    Add-Content $fileName  "</td>"
    Add-Content $fileName  "</tr>"
    Add-Content $fileName  "</table>"
}

Function WriteTableHeader {
    param ($fileName)

    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='10%' align='center'>Drive</td>"
    Add-Content $fileName "<td width='50%' align='center'>Drive Label</td>"
    Add-Content $fileName "<td width='10%' align='center'>Total Capacity (GB)</td>"
    Add-Content $fileName "<td width='10%' align='center'>Used Capacity (GB)</td>"
    Add-Content $fileName "<td width='10%' align='center'>Free Space (GB)</td>"
    Add-Content $fileName "<td width='10%' align='center'>Free Space %</td>"
    Add-Content $fileName "</tr>"
}

Function WriteHTMLFooter {
    param ($fileName)

    Add-Content $fileName "</body>"
    Add-Content $fileName "</html>"
}

Function WriteDiskInfo {
    param ($fileName, $devId, $volName, $frSpace, $totSpace)

    $totSpace = [Math]::Round(($totSpace/1073741824),2)
    $frSpace = [Math]::Round(($frSpace/1073741824),2)
    $usedSpace = $totSpace - $frspace
    $usedSpace = [Math]::Round($usedSpace,2)
    $freePercent = ($frspace/$totSpace)*100
    $freePercent = [Math]::Round($freePercent,0)
 
    if ($volName -eq "Swap_File" ) {$freePercent = "40"}
 
    if ($freePercent -gt $warning) {
        Add-Content $fileName "<tr>"
        Add-Content $fileName "<td>$devid</td>"
        Add-Content $fileName "<td>$volName</td>"
        Add-Content $fileName "<td>$totSpace</td>"
        Add-Content $fileName "<td>$usedSpace</td>"
        Add-Content $fileName "<td>$frSpace</td>"
        Add-Content $fileName "<td bgcolor='#00FF00' align=center>$freePercent</td>"
        Add-Content $fileName "</tr>"
    } elseif ($freePercent -le $critical) {
        Add-Content $fileName "<tr>"
        Add-Content $fileName "<td>$devid</td>"
        Add-Content $fileName "<td>$volName</td>"
        Add-Content $fileName "<td>$totSpace</td>"
        Add-Content $fileName "<td>$usedSpace</td>"
        Add-Content $fileName "<td>$frSpace</td>"
        Add-Content $fileName "<td bgcolor='#FF0000' align=center>$freePercent</td>"
        Add-Content $fileName "</tr>"
    } else {
        Add-Content $fileName "<tr>"
        Add-Content $fileName "<td>$devid</td>"
        Add-Content $fileName "<td>$volName</td>"
        Add-Content $fileName "<td>$totSpace</td>"
        Add-Content $fileName "<td>$usedSpace</td>"
        Add-Content $fileName "<td>$frSpace</td>"
        Add-Content $fileName "<td bgcolor='#FBB917' align=center>$freePercent</td>"
        Add-Content $fileName "</tr>"
    }
}

Function SendEmail {
    param ($from, $to, $subject, $smtpHost, $htmlFileName)

    $body = Get-Content $htmlFileName
    $smtp = New-Object System.Net.Mail.SmtpClient $smtpHost
    $msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
    $msg.isBodyhtml = $true
    $smtp.send($msg)
}

WriteHTMLHeader $freeSpaceFileName

$progress = 0

foreach ($server in $servers) {
    Write-Progress -Activity "Gathering Data..." -Status "$progress out of $($servers.Count) servers." -PercentComplete ($progress/$servers.Count * 100)
    $progress++

    Add-Content $freeSpaceFileName "<table width='100%'><tbody>"
    Add-Content $freeSpaceFileName "<tr bgcolor='#CCCCCC'>"
    if (!$server.StartsWith("#")) {
        Add-Content $freeSpaceFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> $server </strong></font></td>"
        Add-Content $freeSpaceFileName "</tr>"
    } else {
        Add-Content $freeSpaceFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#FF0000' size='2'><strong> $($server.Split("#")[1]) </strong></font></td>"
        Add-Content $freeSpaceFileName "</tr>"
        Add-Content $freeSpaceFileName "</table>"
        continue
    }

    WriteTableHeader $freeSpaceFileName

    try {$disks = Get-WmiObject -ComputerName $server -Class Win32_LogicalDisk -Filter "DriveType='3'" 2>$null} catch {$disks = $null}
    foreach ($disk in $disks) {
        WriteDiskInfo $freeSpaceFileName $disk.DeviceID $disk.VolumeName $disk.FreeSpace $disk.Size
    }

    Add-Content $freeSpaceFileName "</table>"
}
Write-Progress -Activity "Completed" -Completed

WriteHTMLFooter $freeSpaceFileName