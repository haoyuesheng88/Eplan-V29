# Troubleshooting

## No open EPLAN project found

- Confirm that `EPLAN.exe` is running and has a visible main window.
- The resolver accepts several title-bar patterns, including full `.elk` paths and base paths without the extension.
- If the title still cannot be resolved, pass `-ProjectPath` explicitly. This is the recommended fallback on other computers.

## EPLAN installation not found

- The helper first inspects the running `EPLAN.exe` and `W3u.exe` processes.
- If those are missing, it searches common roots under:
  - `C:\Program Files\EPLAN`
  - `C:\Program Files (x86)\EPLAN`
- If your installation is elsewhere, add the correct path by editing `scripts/eplan-common.ps1` or start EPLAN first so the process paths can be discovered.

## Child process returns a non-zero exit code

- This can still be normal with the EPLAN API host.
- If both `eplan_io_points.csv` and `eplan_io_summary.txt` exist, treat the export as successful.

## CSV has empty connection endpoints

- This does not necessarily mean the export failed.
- Many projects expose PLC addresses and function text but not explicit field-side source/target closure through the APIs used here.
- In that case, report that the I/O point list is complete but field-device closure is unavailable from the current project/API path.
