param(
    [string]$InstallDir = "$env:USERPROFILE\mailhandle",
    [switch]$SkipCopy,
    [switch]$SkipLaunch,
    [switch]$SkipCodexLogin
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$pythonPackageId = "Python.Python.3.11"
$nodePackageId = "OpenJS.NodeJS.LTS"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Assert-Command {
    param(
        [string]$Name,
        [string]$HelpMessage
    )

    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw $HelpMessage
    }
}

function Install-WingetPackage {
    param(
        [string]$PackageId,
        [string]$DisplayName
    )

    Write-Step "Installing $DisplayName"
    winget install --id $PackageId -e --accept-package-agreements --accept-source-agreements --silent
}

function Refresh-PathFromMachine {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

Assert-Command -Name "winget" -HelpMessage "winget is required for install_mailhandle.ps1. Install App Installer from Microsoft first."

Install-WingetPackage -PackageId $pythonPackageId -DisplayName "Python 3.11"
Install-WingetPackage -PackageId $nodePackageId -DisplayName "Node.js LTS"
Refresh-PathFromMachine

Assert-Command -Name "py" -HelpMessage "Python launcher 'py' was not found after installation. Open a new terminal and rerun this installer."
Assert-Command -Name "npm" -HelpMessage "npm was not found after Node.js installation. Open a new terminal and rerun this installer."

Write-Step "Installing pywin32"
py -3.11 -m pip install --upgrade pip
py -3.11 -m pip install pywin32

Write-Step "Installing Codex CLI"
npm install -g @openai/codex
Refresh-PathFromMachine

Assert-Command -Name "codex" -HelpMessage "Codex CLI was not found after npm install. Open a new terminal and rerun this installer."

if (-not $SkipCodexLogin) {
    Write-Step "Signing into Codex CLI"
    codex login
}

if (-not $SkipCopy) {
    Write-Step "Copying mailhandle to $InstallDir"
    if (Test-Path $InstallDir) {
        Remove-Item -Path $InstallDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
    Get-ChildItem -Path $projectRoot -Force | Where-Object { $_.Name -notin @(".cache", "data", "dist", "tmp", "__pycache__") } | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination (Join-Path $InstallDir $_.Name) -Recurse -Force
    }
}

$launchRoot = if ($SkipCopy) { $projectRoot } else { $InstallDir }
$launchScript = Join-Path $launchRoot "scripts\launch_mailhandle.ps1"

Write-Host ""
Write-Host "Install complete." -ForegroundColor Green
Write-Host "Mailhandle location: $launchRoot"
Write-Host "Launch command:"
Write-Host "powershell -NoProfile -ExecutionPolicy Bypass -File `"$launchScript`""

if (-not $SkipLaunch) {
    Write-Step "Launching mailhandle"
    powershell -NoProfile -ExecutionPolicy Bypass -File $launchScript
}
