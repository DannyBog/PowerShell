$Broadcast = ([System.Net.IPAddress]::Broadcast)
 
## Create UDP client instance
$UdpClient = New-Object Net.Sockets.UdpClient
 
## Create IP endpoints for each port
$IPEndPoint1 = New-Object Net.IPEndPoint $Broadcast, 0
$IPEndPoint2 = New-Object Net.IPEndPoint $Broadcast, 7
$IPEndPoint3 = New-Object Net.IPEndPoint $Broadcast, 9
 
## Construct physical address instance for the MAC address of the machine (string to byte array)
$MAC = Read-Host "Enter the MAC address of a machine in your LAN"
$MacByteArray = $Mac.Split(":") | ForEach-Object {[Byte] "0x$_"}
 
## Construct the Magic Packet frame
$Packet = [Byte[]](,0xFF*6)+($MacByteArray*16)
 
## Broadcast UDP packets to the IP endpoint of the machine
$UdpClient.Send($Packet, $Packet.Length, $IPEndPoint1) | Out-Null
$UdpClient.Send($Packet, $Packet.Length, $IPEndPoint2) | Out-Null
$UdpClient.Send($Packet, $Packet.Length, $IPEndPoint3) | Out-Null
$UdpClient.Close()