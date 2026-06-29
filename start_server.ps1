# Set current directory to script location
Set-Location $PSScriptRoot

Write-Output "======================================="
Write-Output " Smart Vision CRM - Auto Start Script"
Write-Output "======================================="

# --- Step 1: Start the Node.js server in the background ---
Write-Output "[1/3] Starting CRM backend server..."
$nodeProcess = Start-Process -NoNewWindow -PassThru -FilePath "node" -ArgumentList "server.js" -WorkingDirectory $PSScriptRoot
Write-Output "      Server started (PID: $($nodeProcess.Id))"
Write-Output "      Listening on http://localhost:3000"
Start-Sleep -Seconds 2

# --- Step 2: Start localtunnel and capture the public URL ---
Write-Output "[2/3] Opening public tunnel..."
$tunnelOutput = ""
$tunnelJob = Start-Job -ScriptBlock {
    Set-Location $using:PSScriptRoot
    npx localtunnel --port 3000 --subdomain smartvision-crm 2>&1
}

# Wait up to 15 seconds for the tunnel URL to appear
$timeout = 15
$elapsed = 0
while ($elapsed -lt $timeout) {
    Start-Sleep -Seconds 1
    $elapsed++
    $jobOutput = Receive-Job -Job $tunnelJob -Keep
    $urlLine = $jobOutput | Where-Object { $_ -match "your url is:" }
    if ($urlLine) {
        $tunnelOutput = ($urlLine -replace "your url is: ", "").Trim()
        break
    }
}

if (-not $tunnelOutput) {
    Write-Warning "Could not retrieve tunnel URL. Starting without auto-publish."
    $tunnelOutput = "https://smartvision-crm.loca.lt"
}

Write-Output "      Public URL: $tunnelOutput"

# --- Step 3: Update config.json and push to GitHub ---
Write-Output "[3/3] Publishing tunnel URL to GitHub Pages..."
$configContent = "{ `"backendUrl`": `"$tunnelOutput`" }"
Set-Content -Path "$PSScriptRoot\config.json" -Value $configContent -Encoding UTF8

# Git push
git -C $PSScriptRoot add config.json 2>&1 | Out-Null
git -C $PSScriptRoot commit -m "chore: Update live tunnel backend URL to $tunnelOutput" 2>&1 | Out-Null
git -C $PSScriptRoot push origin main 2>&1 | Out-Null

Write-Output ""
Write-Output "======================================="
Write-Output " ALL SYSTEMS ONLINE!"
Write-Output "======================================="
Write-Output " Local:   http://localhost:3000"
Write-Output " Public:  $tunnelOutput"
Write-Output " Online:  https://smartvision7862.github.io/crm/"
Write-Output "======================================="
Write-Output " Press Ctrl+C to stop the server."
Write-Output "======================================="

# Keep running (wait for the server process to exit)
$nodeProcess.WaitForExit()
