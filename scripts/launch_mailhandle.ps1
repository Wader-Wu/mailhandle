param(
    [ValidateSet("gui", "cli")]
    [string]$Mode = "gui",

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$CliArgs
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$startScript = Join-Path $scriptDir "start_mailhandle.ps1"
$projectRoot = Split-Path -Parent $scriptDir
$summaryFile = Join-Path $projectRoot "tmp\mailhandle-last-start.txt"

if ($Mode -eq "cli") {
    $argumentList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $startScript, "-Mode", "cli")
    if ($CliArgs -and $CliArgs.Count -gt 0) {
        $argumentList += $CliArgs
    }
    & powershell @argumentList
    exit $LASTEXITCODE
}

if (Test-Path $summaryFile) {
    Remove-Item -Path $summaryFile -Force -ErrorAction SilentlyContinue
}

$process = Start-Process powershell `
    -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $startScript, "-Mode", "gui" `
    -PassThru

for ($i = 0; $i -lt 40; $i++) {
    Start-Sleep -Milliseconds 250
    if (Test-Path $summaryFile) {
        Get-Content -Path $summaryFile
        return
    }
    if ($process.HasExited) {
        break
    }
}

if (Test-Path $summaryFile) {
    Get-Content -Path $summaryFile
    return
}

Write-Output "Mailhandle launcher triggered."
Write-Output "PID: $($process.Id)"
Write-Output "Startup summary is not ready yet."
Write-Output "Summary file: $summaryFile"
