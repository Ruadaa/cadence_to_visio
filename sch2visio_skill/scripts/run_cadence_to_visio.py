#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
PROJECT_DIR = SCRIPT_DIR / "project"
CADENCE_ASSET_DIR = SKILL_DIR / "assets" / "cadence"
VISIO_ASSET_DIR = SKILL_DIR / "assets" / "visio"
DEFAULT_STENCIL = VISIO_ASSET_DIR / "circuit.vss"
EXPORTER_FILES = (
    CADENCE_ASSET_DIR / "export_inst_xy_orient.il",
    CADENCE_ASSET_DIR / "export_wire_lines_v4.il",
)


def resolve_path(raw: str | None) -> str | None:
    if raw is None:
        return None
    return str(Path(raw).resolve())


def run_python(script_name: str, args: list[str], cwd: Path | None = None) -> None:
    command = [sys.executable, str(PROJECT_DIR / script_name), *args]
    print("Running:", " ".join(command))
    subprocess.run(command, cwd=str(cwd or PROJECT_DIR), check=True)


def forwarded_args(raw_args: list[str]) -> list[str]:
    if raw_args and raw_args[0] == "--":
        return raw_args[1:]
    return raw_args


def add_common_inputs(command: list[str], args: argparse.Namespace) -> list[str]:
    command.extend(
        [
            "--wires",
            resolve_path(args.wires),
            "--netlist",
            resolve_path(args.netlist),
            "--inst-info",
            resolve_path(args.inst_info),
            "--stencil",
            resolve_path(args.stencil),
        ]
    )
    if args.placement_offsets:
        command.extend(["--placement-offsets", resolve_path(args.placement_offsets)])
    return command


def cmd_paths(_args: argparse.Namespace) -> None:
    print(f"skill_dir={SKILL_DIR}")
    print(f"project_dir={PROJECT_DIR}")
    print(f"stencil={DEFAULT_STENCIL}")
    for exporter in EXPORTER_FILES:
        print(f"exporter={exporter}")


def cmd_copy_exporters(args: argparse.Namespace) -> None:
    destination = Path(args.dest).resolve()
    destination.mkdir(parents=True, exist_ok=True)
    for source in EXPORTER_FILES:
        target = destination / source.name
        shutil.copy2(source, target)
        print(f"Copied {source} -> {target}")


def cmd_validate(args: argparse.Namespace) -> None:
    cwd = Path(args.cwd).resolve()
    cwd.mkdir(parents=True, exist_ok=True)
    command: list[str] = []
    add_common_inputs(command, args)
    command.append("--dry-run")
    command.extend(forwarded_args(args.extra_args))
    run_python("cadence_to_visio_v2.py", command, cwd=cwd)


def cmd_visio(args: argparse.Namespace) -> None:
    cwd = Path(args.cwd).resolve()
    cwd.mkdir(parents=True, exist_ok=True)
    command: list[str] = []
    add_common_inputs(command, args)
    command.extend(forwarded_args(args.extra_args))
    run_python("cadence_to_visio_v2.py", command, cwd=cwd)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Wrapper for the bundled cadence_to_visio project inside the sch2visio skill."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    paths_parser = subparsers.add_parser("paths", help="Print bundled asset locations")
    paths_parser.set_defaults(func=cmd_paths)

    copy_parser = subparsers.add_parser("copy-exporters", help="Copy bundled Cadence exporter scripts to a folder")
    copy_parser.add_argument("--dest", required=True, help="Destination directory for .il exporters")
    copy_parser.set_defaults(func=cmd_copy_exporters)

    for subcommand, handler, help_text in (
        ("validate", cmd_validate, "Validate Visio inputs without opening Microsoft Visio"),
        ("visio", cmd_visio, "Run the Visio reconstruction workflow"),
    ):
        workflow = subparsers.add_parser(subcommand, help=help_text)
        workflow.add_argument("--wires", required=True, help="Path to wires.tsv or wires.xlsx")
        workflow.add_argument("--netlist", required=True, help="Path to CDL netlist.cdl")
        workflow.add_argument("--inst-info", required=True, help="Path to inst_info.txt")
        workflow.add_argument("--stencil", default=str(DEFAULT_STENCIL), help="Path to the Visio stencil (.vss)")
        workflow.add_argument("--placement-offsets", help="Optional placement offset table")
        workflow.add_argument("--cwd", default=".", help="Working directory for generated outputs")
        workflow.add_argument(
            "extra_args",
            nargs=argparse.REMAINDER,
            help="Extra arguments forwarded to cadence_to_visio_v2.py after --",
        )
        workflow.set_defaults(func=handler)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
