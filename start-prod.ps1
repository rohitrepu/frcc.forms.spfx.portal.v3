$ProjectRoot = $PSScriptRoot
$NodePath = Join-Path $ProjectRoot "tools\node"

$env:PATH = "$NodePath;$env:PATH"

Set-Location $ProjectRoot

Write-Host "Project:" $ProjectRoot
Write-Host "Using Node:"
node -v

Write-Host "Using npm:"
npm -v

Write-Host "Building production SPFx package..."

npm run build

Write-Host "Done. Package output should be in:"
Write-Host ".\sharepoint\solution\frcc-forms-portal-v2.sppkg"