---
name: sch2visio
description: Export the currently active Cadence Virtuoso schematic through virtuoso-bridge-lite and generate an editable Microsoft Visio drawing using the bundled Visio reconstruction workflow. Use when the user says "绘制visio", "画visio", "生成Visio", "schematic to Visio", "sch2visio", or asks Codex to export the current Virtuoso schematic into Visio from `netlist.cdl`, `inst_info.txt`, and `wires.tsv`, or wants to validate and draw Visio directly from existing export artifacts on Windows.
---

# Sch2Visio

## Overview

Run the full current-schematic pipeline: execute the remote Virtuoso SKILL exporter, download `netlist.cdl`, `inst_info.txt`, and `wires.tsv`, then call the bundled Visio wrapper to validate and draw the schematic in Visio directly from those files.

## Quick Start

When the user says "绘制visio" with no extra details, run:

```powershell
python <path-to-sch2visio>\scripts\sch2visio.py
```

Default assumptions:

- The target schematic is open and active in Virtuoso.
- `virtuoso-bridge start` is already running and the bridge setup is loaded in CIW.
- A remote `sch2visio.il` exporter entry point exists on the Cadence host.
- Local workspace contains `virtuoso-bridge-lite`, or pass `--bridge-root`.

## Bundled Resources

- Full export-and-draw runner:
  - `scripts/sch2visio.py`
- Direct Visio wrapper for existing artifacts:
  - `scripts/run_cadence_to_visio.py`
- Bundled project code:
  - `scripts/project/cadence_to_visio_v2.py`
  - `scripts/project/cadence_to_visio_core.py`
- Bundled assets:
  - `assets/visio/circuit.vss`

## Common Commands

Use an explicit local output directory:

```powershell
python <path-to-sch2visio>\scripts\sch2visio.py `
  --local-dir C:\path\to\output\SA_PreCh_visio
```

Use a custom remote export directory:

```powershell
python <path-to-sch2visio>\scripts\sch2visio.py `
  --remote-dir /tmp/SA_PreCh_export
```

Forward native `cadence_to_visio_v2.py` options after `--`:

```powershell
python <path-to-sch2visio>\scripts\sch2visio.py `
  -- --skip-nets vdd,vss --skip-mos-body-nets
```

Run export/download/validation without opening Visio:

```powershell
python <path-to-sch2visio>\scripts\sch2visio.py --validate-only
```

Validate existing exported files without running the remote Virtuoso export:

```powershell
python <path-to-sch2visio>\scripts\run_cadence_to_visio.py validate `
  --wires C:\path\wires.tsv `
  --netlist C:\path\netlist.cdl `
  --inst-info C:\path\inst_info.txt `
  --cwd C:\path\output
```

Run the bundled Visio reconstruction directly:

```powershell
python <path-to-sch2visio>\scripts\run_cadence_to_visio.py visio `
  --wires C:\path\wires.tsv `
  --netlist C:\path\netlist.cdl `
  --inst-info C:\path\inst_info.txt `
  --cwd C:\path\output `
  -- --skip-nets vdd,vss --hidden
```

## Outputs

The runner creates these local artifacts:

- `netlist.cdl`: downloaded CDL from Virtuoso.
- `inst_info.txt`: downloaded instance placement/orientation file.
- `wires.tsv`: downloaded wire coordinates.
- `schematic.vsdx`: saved Visio drawing when Visio generation succeeds.

## Troubleshooting

- If SKILL execution blocks on a Virtuoso dialog, run `virtuoso-bridge dismiss-dialog`, then retry.
- If the runner cannot import `virtuoso_bridge`, pass `--bridge-root C:\path\to\virtuoso-bridge-lite`.
- If Visio generation fails after validation, verify Microsoft Visio and `pywin32` are installed.
- If wire validation fails, inspect the downloaded `wires.tsv` header; it must include `group_id`, `seg_id`, `net`, `obj_type`, `layer`, `purpose`, `x1`, `y1`, `x2`, `y2`.
- For exact input contracts and forwarded options, read `references/inputs-and-workflow.md`.
