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

Write-Host "Building production SPFx V4 package..."

npm run build

Write-Host ""
Write-Host "Done. Package output should be in:"
Write-Host ".\sharepoint\solution\frcc-forms-spfx-portal-v4.sppkg"
