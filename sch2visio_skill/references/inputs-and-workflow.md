# Cadence To Visio Inputs

## Required inputs for Visio reconstruction

- `inst_info.txt`: instance name, placement, orientation, bounding box.
- `netlist.cdl`: CDL/SPICE-like netlist for device and net connectivity.
- `wires.tsv`: wire segment coordinates exported from Virtuoso.

The bundled wrapper accepts absolute paths to all three files. Prefer absolute paths when invoking the skill so output files land in the expected working directory.

## Bundled assets

- Cadence exporters:
  - `assets/cadence/export_inst_xy_orient.il`
  - `assets/cadence/export_wire_lines_v4.il`
- Visio stencil:
  - `assets/visio/circuit.vss`

Use `scripts/run_cadence_to_visio.py paths` to print their resolved locations, or `copy-exporters` to copy the `.il` files into a convenient directory before loading them in Virtuoso.

## Typical workflow

1. Export `inst_info.txt` and wire coordinates from Virtuoso with the bundled `.il` files.
2. Save wire coordinates as `wires.tsv`.
3. Save the CDL netlist as `netlist.cdl`.
4. Run a dry validation first:

```powershell
python scripts/run_cadence_to_visio.py validate `
  --wires C:\path\wires.tsv `
  --netlist C:\path\netlist.cdl `
  --inst-info C:\path\inst_info.txt `
  --cwd C:\path\output
```

5. If validation looks good, run the Visio workflow:

```powershell
python scripts/run_cadence_to_visio.py visio `
  --wires C:\path\wires.tsv `
  --netlist C:\path\netlist.cdl `
  --inst-info C:\path\inst_info.txt `
  --cwd C:\path\output `
  -- --skip-nets vdd,vss --hidden
```

## Environment notes

- `visio` and `validate` use the bundled `cadence_to_visio_v2.py`.
- Full Visio reconstruction requires Windows, Microsoft Visio, and `pywin32`.
- `openpyxl` is only needed when you choose to pass `wires.xlsx`; the default TSV path does not require it.

## Common forwarded options

Forward any native project options after `--`:

- `--skip-nets vdd,vss`
- `--skip-mos-body-nets`
- `--draw-mos-b-wires`
- `--wire-adjust snap-endpoints`
- `--hidden`
- `--preserve-absolute`
- `--flip-y`

The wrapper does not reinterpret these flags; it forwards them directly to the bundled project script.
