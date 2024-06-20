$machine = Read-Host "Enter a computer name (or IP)"
$command = Read-Host "Enter a command"

([WMICLASS]"\\$machine\ROOT\CIMV2:win32_process").Create("cmd.exe `/c $command")