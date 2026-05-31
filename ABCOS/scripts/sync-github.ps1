# abcos 작업 폴더 → seobju-github/ABCOS 동기화 후 GitHub push
param(
    [switch]$Push,
    [string]$Message = ""
)

$ErrorActionPreference = "Stop"
$git = "C:\Program Files\Git\cmd\git.exe"
if (-not (Test-Path $git)) {
    throw "Git not found. Install from https://git-scm.com/download/win"
}

$abcos = Split-Path $PSScriptRoot -Parent
$repoRoot = Join-Path (Split-Path $abcos -Parent) "seobju-github"
$dest = Join-Path $repoRoot "ABCOS"

if (-not (Test-Path $repoRoot)) {
    throw "GitHub clone folder not found: $repoRoot"
}

Write-Host "Sync: $abcos -> $dest"

$excludeDirs = @(".git", "node_modules", "__pycache__", ".venv", ".vercel", ".cache", ".local", ".config")
$excludeArgs = $excludeDirs | ForEach-Object { "/XD", $_ }

& robocopy $abcos $dest /E /NFL /NDL /NJH /NJS @excludeArgs /XF "*.pyc" | Out-Null
$rc = $LASTEXITCODE
if ($rc -ge 8) {
    throw "robocopy failed with exit code $rc"
}

Push-Location $repoRoot
try {
    & $git add ABCOS README.md
    $status = & $git status --porcelain
    if (-not $status) {
        Write-Host "No changes to commit in seobju-github."
        if ($Push) {
            & $git push origin main
        }
        return
    }

    & $git status --short
    if (-not $Message) {
        $Message = Read-Host "Commit message"
    }
    if (-not $Message) {
        $Message = "Sync from abcos workspace"
    }

    & $git commit -m $Message
    Write-Host "Committed: $Message"

    if ($Push) {
        & $git push origin main
        Write-Host "Pushed to https://github.com/irismonapark/seobju.git"
    } else {
        Write-Host "Local commit done. Push with: .\scripts\sync-github.ps1 -Push"
    }
}
finally {
    Pop-Location
}
