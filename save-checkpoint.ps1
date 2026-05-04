param(
  [string]$Message = "Auto checkpoint"
)

Set-Location "C:\Users\S03249519\Projects\frcc.forms.spfx.portal.v2"

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$status = git status --porcelain

if ([string]::IsNullOrWhiteSpace($status)) {
  Write-Host "No changes to commit at $timestamp"
  exit 0
}

git add .
git commit -m "$Message - $timestamp"

Write-Host "Checkpoint committed at $timestamp"
