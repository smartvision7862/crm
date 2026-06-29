# Set current directory to script location
Set-Location $PSScriptRoot

Write-Output "Starting WhatsApp CRM Integration Server..."
Write-Output "Listening on http://localhost:3000"

# Use global node if available, otherwise fallback to local node bin
if (Get-Command node -ErrorAction SilentlyContinue) {
    node server.js
} else {
    $localNode = "$PSScriptRoot\node_bin\node-v20.12.2-win-x64\node.exe"
    if (Test-Path $localNode) {
        & $localNode server.js
    } else {
        Write-Error "Node.js was not found globally, and local fallback binary was not found."
    }
}
