$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$runScript = Join-Path $scriptDir "run_mail_database.py"
$tmpDir = Join-Path $projectRoot "tmp"
$envFile = Join-Path $scriptDir ".env"
$pidFile = Join-Path $tmpDir "mailhandle-last.pid"

New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

function Get-EnvValue {
    param(
        [string]$Path,
        [string]$Name
    )

    if (-not (Test-Path $Path)) {
        return $null
    }

    foreach ($line in Get-Content $Path) {
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed.StartsWith("#")) {
            continue
        }
        if ($trimmed -match "^\Q$Name\E\s*=\s*(.*)$") {
            $value = $Matches[1].Trim()
            if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }
            return $value
        }
    }

    return $null
}

$pythonExe = Get-EnvValue -Path $envFile -Name "WINDOWS_PYTHON_EXE"
if (-not $pythonExe) {
    $pythonExe = "python"
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$stdoutLog = Join-Path $tmpDir "mailhandle-start-$timestamp.stdout.log"
$stderrLog = Join-Path $tmpDir "mailhandle-start-$timestamp.stderr.log"
$summaryFile = Join-Path $tmpDir "mailhandle-last-start.txt"
$summaryLines = @()
$exitCode = 0

try {
    if (Test-Path $pidFile) {
        $previousPid = (Get-Content $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
        if ($previousPid -match "^\d+$") {
            try {
                Stop-Process -Id ([int]$previousPid) -Force -ErrorAction Stop
            } catch {
            }
        }
    }

    $process = Start-Process `
        -FilePath $pythonExe `
        -ArgumentList "-u", $runScript, "--open-browser" `
        -WorkingDirectory $projectRoot `
        -PassThru `
        -RedirectStandardOutput $stdoutLog `
        -RedirectStandardError $stderrLog

    Set-Content -Path $pidFile -Value "$($process.Id)" -Encoding Ascii

    $summaryLines = @(
        "Mailhandle launcher started."
        "PID: $($process.Id)"
        "Stdout: $stdoutLog"
        "Stderr: $stderrLog"
        "Browser auto-open requested."
    )

    $workspaceLine = $null
    for ($i = 0; $i -lt 10; $i++) {
        Start-Sleep -Milliseconds 250
        if (-not (Test-Path $stdoutLog)) {
            continue
        }
        $workspaceMatch = Select-String -Path $stdoutLog -Pattern "^Mailhandle workspace:" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($workspaceMatch) {
            $workspaceLine = $workspaceMatch.Line
            break
        }
        if ($process.HasExited) {
            break
        }
    }

    if ($workspaceLine) {
        $summaryLines += $workspaceLine
        $summaryLines += "Workspace startup confirmed."
    } elseif ($process.HasExited) {
        $summaryLines += "Workspace process exited before publishing the URL. Check the stderr log."
        $exitCode = 1
    } else {
        $summaryLines += "Workspace is still starting. Check the stdout log for the URL."
    }
} catch {
    $summaryLines = @(
        "Mailhandle launcher failed."
        "Details: $($_.Exception.Message)"
        "Stdout: $stdoutLog"
        "Stderr: $stderrLog"
    )
    $exitCode = 1
}

$summaryLines | Set-Content -Path $summaryFile -Encoding Ascii
$summaryLines | ForEach-Object { Write-Output $_ }
exit $exitCode
