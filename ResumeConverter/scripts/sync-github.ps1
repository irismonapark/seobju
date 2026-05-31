# ResumeConverter → seobju-github/ResumeConverter 동기화 후 GitHub push
param(
    [switch]$Push,
    [string]$Message = ""
)

$ErrorActionPreference = "Stop"
$git = "C:\Program Files\Git\bin\git.exe"
if (-not (Test-Path $git)) {
    $git = "C:\Program Files\Git\cmd\git.exe"
}
if (-not (Test-Path $git)) {
    throw "Git not found. Install from https://git-scm.com/download/win"
}

$converter = Split-Path $PSScriptRoot -Parent
$repoRoot = Join-Path (Split-Path $converter -Parent) "seobju-github"
$dest = Join-Path $repoRoot "ResumeConverter"

if (-not (Test-Path $repoRoot)) {
    throw "GitHub clone folder not found: $repoRoot`nRun: git clone https://github.com/irismonapark/seobju.git `"$repoRoot`""
}

Write-Host "Sync: $converter -> $dest"

$excludeDirs = @(".git", "__pycache__", ".pythonlibs", ".vercel", ".cursor")
$excludeArgs = $excludeDirs | ForEach-Object { "/XD", $_ }

& robocopy $converter $dest /E /NFL /NDL /NJH /NJS @excludeArgs /XF "*.pyc" | Out-Null
if ($LASTEXITCODE -ge 8) {
    throw "robocopy failed with exit code $LASTEXITCODE"
}

Push-Location $repoRoot
try {
    & $git pull origin main
    & $git add ResumeConverter README.md
    $status = & $git status --porcelain
    if (-not $status) {
        Write-Host "No changes to commit in seobju-github."
        if ($Push) {
            & $git push origin main
            Write-Host "Pushed (up to date): https://github.com/irismonapark/seobju"
        }
        return
    }

    & $git status --short
    if (-not $Message) {
        $Message = Read-Host "Commit message"
    }
    if (-not $Message) {
        $Message = "Sync ResumeConverter from workspace"
    }

    & $git commit -m $Message
    Write-Host "Committed: $Message"

    if ($Push) {
        & $git push origin main
        Write-Host "Pushed to https://github.com/irismonapark/seobju.git"
    }
}
finally {
    Pop-Location
}
