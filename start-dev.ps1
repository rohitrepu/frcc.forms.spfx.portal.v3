$ProjectRoot = $PSScriptRoot
$NodePath = Join-Path $ProjectRoot "tools\node"

$env:PATH = "$NodePath;$env:PATH"

Set-Location $ProjectRoot

Write-Host "Project:" $ProjectRoot
Write-Host "Using Node:"
node -v

Write-Host "Using npm:"
npm -v

if (!(Test-Path ".\node_modules")) {
    Write-Host "node_modules missing. Installing dependencies..."
    npm install
}

$Port = 4321
$PortInUse = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue

if ($PortInUse) {
    Write-Host "Port 4321 is already in use. Closing existing process..."
    $PortInUse | ForEach-Object {
        Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
    }
}

$DebugUrl = "https://cccs.sharepoint.com/sites/frcc-forms/SitePages/FRCC-Forms-Portal.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true"
$ManifestUrl = "https://localhost:4321/temp/build/manifests.js"

Write-Host "Starting SPFx dev server..."

$Job = Start-Job -ScriptBlock {
    param($ProjectRoot, $NodePath)

    $env:PATH = "$NodePath;$env:PATH"
    Set-Location $ProjectRoot
    npm run start
} -ArgumentList $ProjectRoot, $NodePath

Write-Host "Waiting for SPFx manifest to become available..."

$Ready = $false

for ($i = 1; $i -le 60; $i++) {
    try {
        Invoke-WebRequest -Uri $ManifestUrl -UseBasicParsing -SkipCertificateCheck -TimeoutSec 2 | Out-Null
        $Ready = $true
        break
    }
    catch {
        Start-Sleep -Seconds 2
    }
}

if ($Ready) {
    Write-Host "SPFx is ready. Opening SharePoint debug page..."
    Start-Process $DebugUrl
}
else {
    Write-Host "SPFx did not become ready automatically. Open this URL manually after the server finishes loading:"
    Write-Host $DebugUrl
}

Write-Host "SPFx server is running. Keep this PowerShell window open."
Receive-Job $Job -Wait