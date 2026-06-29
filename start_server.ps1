$nodeDir = "c:\Users\uniqu\Downloads\CRM\node_bin\node-v20.12.2-win-x64"
$env:PATH = "$nodeDir;" + $env:PATH

Set-Location "c:\Users\uniqu\Downloads\CRM"

Write-Output "Starting WhatsApp CRM Integration Server..."
Write-Output "Listening on http://localhost:3000"
& "$nodeDir\node.exe" server.js
