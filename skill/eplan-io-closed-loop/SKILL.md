---
name: eplan-io-closed-loop
description: Connect to an already open EPLAN Electric P8 project on Windows, discover the active project, export a read-only PLC IO point list, and summarize the closed-loop input/output map as CSV and text. Use when the user asks to connect to EPLAN, Electric P8, a live EPLAN project, or wants PLC IO points, IO closure, module/address listings, or a reusable export workflow.
---

# EPLAN IO Closed Loop

Use this skill to turn a live EPLAN Electric P8 session into a repeatable read-only PLC I/O export workflow.

Assume Windows, PowerShell, and an already open EPLAN project. Prefer the offline DataModel API for extraction. Do not edit the live project unless the user explicitly asks for project changes.

## Workflow

1. Confirm that an `EPLAN` process is open or that the user gave a valid `.elk` path.
2. Resolve the active project path from the EPLAN main window title when the path is not supplied.
3. If title-bar parsing is ambiguous or the title contains only a project name, ask for or use an explicit `.elk` path.
4. Resolve the local EPLAN installation dynamically instead of assuming one fixed versioned path.
5. Run the export in a child `powershell.exe -STA` process.
6. Produce `eplan_io_points.csv` and `eplan_io_summary.txt` in the requested output directory.
7. Report module counts, direction counts, and a few representative points.

## Primary Scripts

- `scripts/run-open-project-io-export.ps1`
  Main entrypoint. Use this when the user wants the current EPLAN project's I/O list.
- `scripts/find-open-eplan-project.ps1`
  Resolve the active `.elk` file from the open EPLAN window title when the title exposes a path-like segment.
- `scripts/export-eplan-io.ps1`
  Open the project in read-only mode through the EPLAN offline API and export PLC I/O points.
- `scripts/eplan-common.ps1`
  Shared helpers for EPLAN installation detection and project-path resolution.

## Default Command

When the user simply wants the current project's I/O list:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\skill\eplan-io-closed-loop\scripts\run-open-project-io-export.ps1" -OutputDir "C:\path\to\output"
```

If the project path is already known:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\skill\eplan-io-closed-loop\scripts\run-open-project-io-export.ps1" -ProjectPath "C:\Projects\Example.elk" -OutputDir "C:\path\to\output"
```

## Output Files

- `eplan_io_points.csv`
  One row per PLC I/O point with module kind, PLC box, terminal, address, symbolic address, data type, function text, and page.
- `eplan_io_summary.txt`
  Project path, CSV path, PLC box count, I/O point count, direction counts, and module-kind counts.
- `eplan_io_error.txt`
  Present only when the export throws.

## Important Rules

- Keep the extraction read-only. Use `ProjectManager.OpenMode.ReadOnly`.
- Set `LockProjectByDefault = $false` before opening the project.
- Detect the EPLAN installation dynamically. Prefer the running `EPLAN.exe` and `W3u.exe` paths before searching common installation roots.
- Treat explicit `-ProjectPath` input as the most reliable option when moving the workflow between different computers.
- Use a child `powershell.exe -STA` host for the export. This keeps the EPLAN API session isolated and avoids fragile host-state issues.
- Treat the export as successful when the CSV and summary exist, even if the child process returns a non-zero exit code.
- Prefer environment variables or explicit parameters over hardcoding Chinese project paths inside `.ps1` source.

## Interpretation Rules

- Infer `Direction` from the PLC box prefix when EPLAN leaves `PLCIOENTRY_DIRECTION` empty:
  - `AI`, `DI` => `Input`
  - `AQ`, `DQ` => `Output`
- Trim internal formatting prefixes from `FunctionText` by keeping only the content after the last `@` and removing the trailing `;`.
- Treat `ConnectionEndpoints` and `ExternalConnections` as optional enrichment only. Many projects expose the PLC addresses and comments but not explicit field-side closure through this API path.

## What To Report

Include:

- project path
- generated file paths
- PLC box count
- I/O point count
- counts by `AI`, `AQ`, `DI`, `DQ`
- counts by `Input` and `Output`
- a few representative points such as `I0.0`, `Q0.0`, `IW0+`, or `QW0+`

If the user asks for field-device closure and the CSV has empty `ConnectionEndpoints`, say that the module/address/function loop was exported successfully, but field-side source/target closure is not exposed by this project through the current API route.

Read `references/troubleshooting.md` when the export fails or when EPLAN is installed in an unusual location.
