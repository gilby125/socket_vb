$ipaddress = "4.2.2.1"
$port = 53
$connection = New-Object System.Net.Sockets.TcpClient($ipaddress, $port)

if ($connection.Connected) {
    Write-Host "Success"
}
else {
    Write-Host "Failed"
}
