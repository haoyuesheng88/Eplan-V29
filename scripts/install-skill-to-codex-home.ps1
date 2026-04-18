param(
    [string]$TargetRoot
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceSkill = Join-Path $repoRoot 'skill\eplan-io-closed-loop'

if (-not (Test-Path -LiteralPath $sourceSkill)) {
    throw "Skill source was not found: $sourceSkill"
}

if ([string]::IsNullOrWhiteSpace($TargetRoot)) {
    if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
        $TargetRoot = Join-Path $env:CODEX_HOME 'skills'
    }
    else {
        $TargetRoot = Join-Path (Join-Path $HOME '.codex') 'skills'
    }
}

New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null

$targetSkill = Join-Path $TargetRoot 'eplan-io-closed-loop'
$backupPath = $null

if (Test-Path -LiteralPath $targetSkill) {
    $backupPath = '{0}.bak-{1}' -f $targetSkill, (Get-Date -Format 'yyyyMMdd-HHmmss')
    Move-Item -LiteralPath $targetSkill -Destination $backupPath
}

Copy-Item -LiteralPath $sourceSkill -Destination $targetSkill -Recurse -Force

[pscustomobject]@{
    SourceSkill = $sourceSkill
    InstalledSkill = $targetSkill
    BackupPath = $backupPath
}
