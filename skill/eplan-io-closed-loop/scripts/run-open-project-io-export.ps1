param(
    [string]$ProjectPath,
    [string]$OutputDir = (Join-Path (Get-Location) 'output')
)

$ErrorActionPreference = 'Stop'

$exportScript = Join-Path $PSScriptRoot 'export-eplan-io.ps1'

if (-not (Test-Path -LiteralPath $exportScript)) {
    throw "Export script was not found: $exportScript"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$invokeArgs = @(
    '-STA',
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', $exportScript,
    '-OutputDir', $OutputDir
)

if (-not [string]::IsNullOrWhiteSpace($ProjectPath)) {
    $invokeArgs += @('-ProjectPath', $ProjectPath)
}

& powershell.exe @invokeArgs
$childExitCode = $LASTEXITCODE

$summaryPath = Join-Path $OutputDir 'eplan_io_summary.txt'
$csvPath = Join-Path $OutputDir 'eplan_io_points.csv'
$errorPath = Join-Path $OutputDir 'eplan_io_error.txt'

if ((Test-Path -LiteralPath $summaryPath) -and (Test-Path -LiteralPath $csvPath)) {
    [pscustomobject]@{
        ProjectPath = if ($ProjectPath) { $ProjectPath } else { '' }
        OutputDir = $OutputDir
        SummaryPath = $summaryPath
        CsvPath = $csvPath
        ChildExitCode = $childExitCode
    }
    exit 0
}

if (Test-Path -LiteralPath $errorPath) {
    throw (Get-Content -LiteralPath $errorPath -Raw)
}

throw "EPLAN IO export did not produce output files. ChildExitCode=$childExitCode"
