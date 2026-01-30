# CTI Dashboard Web Server
# Run this to start the dashboard server

$port = 3000
$dashboardPath = $PSScriptRoot

Write-Host "Starting CTI Dashboard Server on port $port..."
Write-Host "Dashboard URL: http://localhost:$port/cti-dashboard-live.html"
Write-Host ""
Write-Host "To access from other computers on your network:"
$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch "Loopback" -and $_.IPAddress -notmatch "^169" } | Select-Object -First 1).IPAddress
Write-Host "  http://${ip}:$port/cti-dashboard-live.html"
Write-Host ""
Write-Host "Press Ctrl+C to stop the server"

# Start serve
serve -p $port -s $dashboardPath
