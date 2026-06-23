Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repo

$masterDir = Join-Path $repo 'backups/master_snapshot'
$latestDir = Join-Path $repo 'backups/latest_snapshot'

if (-not (Test-Path $masterDir)) {
    throw "Missing master snapshot folder: $masterDir"
}

$publishFiles = @(
    'tidplan_gantt_status_per_land.html',
    'tidplan_gantt_kalkyl.html',
    'index.html',
    'tidplan_gantt_status_per_land.md',
    'tidplan_gantt_generated.md'
)

# Keep master snapshot in sync with current workspace so publish always uses latest updates.
foreach ($file in $publishFiles) {
    $workspaceSource = Join-Path $repo $file
    $masterTarget = Join-Path $masterDir $file
    if (Test-Path $workspaceSource) {
        Copy-Item -Path $workspaceSource -Destination $masterTarget -Force
    }
}

$copied = @()
foreach ($file in $publishFiles) {
    $source = Join-Path $masterDir $file
    $target = Join-Path $repo $file
    if (Test-Path $source) {
        Copy-Item -Path $source -Destination $target -Force
        $copied += $file
    }
}

if (-not $copied.Count) {
    throw 'No publishable files were found in backups/master_snapshot.'
}

if (-not (Test-Path $latestDir)) {
    New-Item -ItemType Directory -Path $latestDir -Force | Out-Null
}

foreach ($file in $publishFiles) {
    $source = Join-Path $repo $file
    $target = Join-Path $latestDir $file
    if (Test-Path $source) {
        Copy-Item -Path $source -Destination $target -Force
    }
}

$latestStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Set-Content -Path (Join-Path $repo 'backups/LATEST_BACKUP.txt') -Value $latestStamp -Encoding UTF8

git add -- $copied

$pending = git diff --cached --name-only
if (-not $pending) {
    Write-Output 'No published file changes detected. Nothing to commit.'
    exit 0
}

$commitMessage = "Publish master snapshot $latestStamp"
git commit -m $commitMessage | Out-Null
git push origin main | Out-Null

Write-Output "Published files to GitHub Pages: $($copied -join ', ')"
Write-Output "Commit: $commitMessage"