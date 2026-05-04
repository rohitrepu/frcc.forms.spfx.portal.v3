Set-Location "$PSScriptRoot"

$today = Get-Date -Format "yyyy-MM-dd"
$stampFile = ".\.vscode\.last-auto-backup"

if (Test-Path $stampFile) {
    $lastRun = Get-Content $stampFile -Raw
    if ($lastRun.Trim() -eq $today) {
        Write-Host "Auto-backup already ran today."
        exit 0
    }
}

$status = git status --porcelain

if ([string]::IsNullOrWhiteSpace($status)) {
    Write-Host "No changes to back up."
    Set-Content -Path $stampFile -Value $today
    exit 0
}

git add .
git commit -m "V3 backup: $today"
git push origin main
git push personal-backup main

Set-Content -Path $stampFile -Value $today

Write-Host "Auto-backup complete for $today."