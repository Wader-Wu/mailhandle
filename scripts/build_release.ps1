param(
    [string]$Version = "",
    [string]$PortablePythonDir = "",
    [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir

if (-not $Version) {
    $Version = Get-Date -Format "yyyy.MM.dd"
}
if (-not $OutputDir) {
    $OutputDir = Join-Path $projectRoot "dist"
}

$sourceStage = Join-Path $OutputDir "mailhandle-$Version-source"
$portableStage = Join-Path $OutputDir "mailhandle-$Version-windows-portable"
$sourceZip = "$sourceStage.zip"
$portableZip = "$portableStage.zip"

$topLevelIncludes = @(
    "README.md",
    "SKILL.md",
    "start_mailhandle_gui.bat"
)

$treeIncludes = @(
    "references",
    "scripts"
)

$skipDirNames = @(
    ".cache",
    "data",
    "dist",
    "log",
    "records",
    "sessions",
    "tmp",
    "__pycache__"
)

$skipFileNames = @(
    "auth.json",
    "history.jsonl",
    "models_cache.json",
    "priority_rules.json.bak"
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Reset-StageDir {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force
    }
    New-Item -ItemType Directory -Path $Path | Out-Null
}

function Copy-TreeFiltered {
    param(
        [string]$SourceRoot,
        [string]$RelativePath,
        [string]$DestinationRoot
    )

    $sourcePath = Join-Path $SourceRoot $RelativePath
    if (-not (Test-Path $sourcePath)) {
        return
    }

    Get-ChildItem -Path $sourcePath -Recurse -Force | ForEach-Object {
        $fullPath = $_.FullName
        $relative = $fullPath.Substring($SourceRoot.Length).TrimStart('\')
        $parts = $relative -split '[\\/]'

        if ($_.PSIsContainer) {
            if ($parts | Where-Object { $skipDirNames -contains $_ }) {
                return
            }
            $targetDir = Join-Path $DestinationRoot $relative
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir | Out-Null
            }
            return
        }

        if ($parts | Where-Object { $skipDirNames -contains $_ }) {
            return
        }

        $fileName = $_.Name
        if ($skipFileNames -contains $fileName) {
            return
        }
        if ($fileName -like "*.pyc") {
            return
        }
        if ($fileName -eq ".env") {
            return
        }

        $targetPath = Join-Path $DestinationRoot $relative
        $targetParent = Split-Path -Parent $targetPath
        if (-not (Test-Path $targetParent)) {
            New-Item -ItemType Directory -Path $targetParent | Out-Null
        }
        Copy-Item -Path $fullPath -Destination $targetPath -Force
    }
}

function Write-ZipFromDirectory {
    param(
        [string]$SourceDir,
        [string]$DestinationZip
    )

    if (Test-Path $DestinationZip) {
        Remove-Item -Path $DestinationZip -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDir, $DestinationZip, [System.IO.Compression.CompressionLevel]::Optimal, $false)
}

Reset-StageDir -Path $sourceStage

foreach ($item in $topLevelIncludes) {
    $sourcePath = Join-Path $projectRoot $item
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination (Join-Path $sourceStage $item) -Force
    }
}

foreach ($tree in $treeIncludes) {
    Copy-TreeFiltered -SourceRoot $projectRoot -RelativePath $tree -DestinationRoot $sourceStage
}

Write-ZipFromDirectory -SourceDir $sourceStage -DestinationZip $sourceZip

$summary = @(
    "Source package: $sourceZip"
)

if ($PortablePythonDir) {
    $portablePythonExe = Join-Path $PortablePythonDir "python.exe"
    if (-not (Test-Path $portablePythonExe)) {
        throw "PortablePythonDir must contain python.exe. Received: $PortablePythonDir"
    }

    Reset-StageDir -Path $portableStage
    Copy-Item -Path (Join-Path $sourceStage "*") -Destination $portableStage -Recurse -Force

    $runtimeTarget = Join-Path $portableStage "runtime\python"
    New-Item -ItemType Directory -Force -Path $runtimeTarget | Out-Null
    Copy-Item -Path (Join-Path $PortablePythonDir "*") -Destination $runtimeTarget -Recurse -Force

    Write-ZipFromDirectory -SourceDir $portableStage -DestinationZip $portableZip
    $summary += "Portable Windows package: $portableZip"
    $summary += "Bundled runtime: $runtimeTarget"
}

$summary += ""
$summary += "Next step for a truly easy Windows release:"
$summary += "- ship the portable zip with runtime\\python already populated"
$summary += "- optional LLM features still need Codex CLI today"

$summary | ForEach-Object { Write-Output $_ }
