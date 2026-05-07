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

$Port = 5432

$PortInUse = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue

if ($PortInUse) {
    Write-Host "Port $Port is already in use. Closing existing process..."

    $PortInUse | ForEach-Object {
        Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
    }

    Start-Sleep -Seconds 2
}

$DebugUrl = "https://cccs.sharepoint.com/sites/frcc-forms/_layouts/15/workbench.aspx?debugManifestsFile=https://localhost:5432/temp/build/manifests.js&debug=true&noredir=true"

$ManifestUrl = "https://localhost:5432/temp/build/manifests.js"

Write-Host "Starting SPFx dev server on port $Port..."

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
        Invoke-WebRequest `
            -Uri $ManifestUrl `
            -UseBasicParsing `
            -SkipCertificateCheck `
            -TimeoutSec 2 | Out-Null

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
    Write-Host ""
    Write-Host "SPFx did not become ready automatically."
    Write-Host "Open this URL manually after webpack finishes:"
    Write-Host ""
    Write-Host $DebugUrl
}

Write-Host ""
Write-Host "SPFx server is running."
Write-Host "Keep this PowerShell window open."

Receive-Job $Job -Wait