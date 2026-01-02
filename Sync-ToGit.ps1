<#
.SYNOPSIS
    Syncs files from Wrestling Robe working folder to Git repo and pushes to GitHub.

.DESCRIPTION
    1. Copies specified files from working folder to git repo
    2. Stages all changes
    3. Commits with user-provided message
    4. Pushes to GitHub

.PARAMETER CommitMessage
    Optional commit message. If not provided, prompts for one.

.PARAMETER Branch
    Target branch (default: Main). Use Research-Corpus-Generation or Document-Image-Analysis for other branches.

.EXAMPLE
    .\Sync-ToGit.ps1
    .\Sync-ToGit.ps1 -CommitMessage "Update ScholarSweep"
    .\Sync-ToGit.ps1 -Branch "Research-Corpus-Generation"
#>

param(
    [string]$CommitMessage = "",
    [string]$Branch = "Main"
)

# Paths
$WorkingFolder = "C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI"
$GitRepo = "C:\Users\azt12\OneDrive\Documents\Git Wrestling Robe"

# Files to sync per branch
$BranchFiles = @{
    "Main" = @(
        "CLAUDE.md"
    )
    "Research-Corpus-Generation" = @(
        "ScholarSweep.py"
    )
    "Document-Image-Analysis" = @(
        "TextileVision_v01.1.py",
        "TextileVision_v01.2.py",
        "TextileVision_v01.3.py",
        "FabricETL.py",
        "ChartOCRTester.py",
        "F2TBench.py"
    )
}

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  FabricETL Git Sync Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan

# Change to git repo
Set-Location $GitRepo

# Checkout branch
Write-Host "`nSwitching to branch: $Branch" -ForegroundColor Yellow
git checkout $Branch 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Branch '$Branch' not found locally. Fetching..." -ForegroundColor Yellow
    git fetch origin $Branch
    git checkout -b $Branch origin/$Branch
}

git pull origin $Branch

# Get files for this branch
$FilesToSync = $BranchFiles[$Branch]
if (-not $FilesToSync) {
    Write-Host "No files configured for branch '$Branch'" -ForegroundColor Red
    exit 1
}

Write-Host "`nFiles to sync for $Branch branch:" -ForegroundColor Yellow
$FilesToSync | ForEach-Object { Write-Host "  - $_" }

# Copy files
Write-Host "`nCopying files..." -ForegroundColor Yellow
$CopiedCount = 0
foreach ($File in $FilesToSync) {
    $Source = Join-Path $WorkingFolder $File
    $Dest = Join-Path $GitRepo $File

    if (Test-Path $Source) {
        Copy-Item $Source $Dest -Force
        Write-Host "  Copied: $File" -ForegroundColor Green
        $CopiedCount++
    } else {
        Write-Host "  NOT FOUND: $File" -ForegroundColor Red
    }
}

if ($CopiedCount -eq 0) {
    Write-Host "`nNo files copied. Exiting." -ForegroundColor Red
    exit 1
}

# Check for changes
$Status = git status --porcelain
if (-not $Status) {
    Write-Host "`nNo changes detected. Nothing to commit." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nChanges detected:" -ForegroundColor Yellow
git status --short

# Get commit message
if (-not $CommitMessage) {
    Write-Host "`nEnter commit message (or press Enter for default):" -ForegroundColor Cyan
    $CommitMessage = Read-Host
    if (-not $CommitMessage) {
        $CommitMessage = "Update files from working folder"
    }
}

# Stage, commit, push
Write-Host "`nStaging changes..." -ForegroundColor Yellow
git add -A

Write-Host "Committing..." -ForegroundColor Yellow
$FullMessage = @"
$CommitMessage

Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
"@

git commit -m $FullMessage

Write-Host "Pushing to origin/$Branch..." -ForegroundColor Yellow
git push origin $Branch

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n======================================" -ForegroundColor Green
    Write-Host "  SUCCESS! Changes pushed to GitHub" -ForegroundColor Green
    Write-Host "======================================" -ForegroundColor Green
} else {
    Write-Host "`nPush failed. Check error above." -ForegroundColor Red
}
