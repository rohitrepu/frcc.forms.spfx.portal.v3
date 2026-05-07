Set-Location "C:\Users\S03249519\Projects\frcc.forms.spfx.portal.v4"

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$status = git status --porcelain

if ([string]::IsNullOrWhiteSpace($status)) {
  Write-Host "No V4 changes to back up at $timestamp"
  exit 0
}

git add .
git commit -m "Auto backup V4 - $timestamp"

git push origin main
git push personal-backup main
git push work-backup main

Write-Host "V4 auto backup pushed to all remotes at $timestamp"
