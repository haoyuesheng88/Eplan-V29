param(
    [string]$ProjectPath,
    [string]$OutputDir = (Join-Path (Get-Location) 'output')
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$runner = Join-Path $repoRoot 'skill\eplan-io-closed-loop\scripts\run-open-project-io-export.ps1'

if (-not (Test-Path -LiteralPath $runner)) {
    throw "Skill runner was not found: $runner"
}

$invokeParams = @{
    OutputDir = $OutputDir
}

if (-not [string]::IsNullOrWhiteSpace($ProjectPath)) {
    $invokeParams.ProjectPath = $ProjectPath
}

& $runner @invokeParams
