# Eplan-V29

Portable Codex skill and helper scripts for exporting PLC I/O points from an open EPLAN Electric P8 project on Windows.

## What Is Included

- `skill/eplan-io-closed-loop/`
  The reusable Codex skill.
- `scripts/install-skill-to-codex-home.ps1`
  Copy the skill into the local Codex skills directory on another computer.
- `scripts/export-open-eplan-io.ps1`
  Run the repository copy directly without installing the skill first.

## Quick Start

Install the skill into the local Codex skills directory:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-skill-to-codex-home.ps1
```

Run the export directly from the repository:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\export-open-eplan-io.ps1 -OutputDir .\output
```

If the EPLAN project is already known, pass it explicitly:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\export-open-eplan-io.ps1 -ProjectPath "C:\Projects\Example.elk" -OutputDir .\output
```

If auto-detection cannot resolve the current project from the EPLAN window title, `-ProjectPath` is the recommended fallback and is the most portable way to run the export across different machines and UI layouts.

## Output Files

- `eplan_io_points.csv`
  One row per PLC I/O point.
- `eplan_io_summary.txt`
  Project-level totals and module/direction counts.
- `eplan_io_error.txt`
  Present only if the export failed.

## Notes

- The workflow is read-only and does not modify the EPLAN project.
- The scripts auto-detect common EPLAN Electric P8 installations instead of hardcoding one computer's paths.
- The project-path resolver accepts several EPLAN title-bar formats, but some installations show only a project name instead of a full path.
- Non-zero child PowerShell exit codes can still be normal if the CSV and summary were written successfully.
