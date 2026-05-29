#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
CADENCE_TO_VISIO_WRAPPER = SCRIPT_DIR / "run_cadence_to_visio.py"

DEFAULT_REMOTE_SKILL = os.environ.get("SCH2VISIO_REMOTE_SKILL", "/tmp/sch2visio.il")
DEFAULT_REMOTE_BASE = os.environ.get("SCH2VISIO_REMOTE_BASE", "/tmp")
DEFAULT_LOCAL_BASE = Path.cwd() / "output"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export active Virtuoso schematic and generate a Visio drawing."
    )
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    parser.add_argument(
        "--bridge-root",
        help="Path to virtuoso-bridge-lite. Defaults to ./virtuoso-bridge-lite when present.",
    )
    parser.add_argument(
        "--remote-skill",
        default=DEFAULT_REMOTE_SKILL,
        help="Remote SKILL file that defines sch2visio(outDir).",
    )
    parser.add_argument(
        "--remote-dir",
        default=f"{DEFAULT_REMOTE_BASE}/sch2visio_export_{stamp}",
        help="Remote directory where sch2visio.il writes netlist.cdl, inst_info.txt, and wires.tsv.",
    )
    parser.add_argument(
        "--local-dir",
        default=str(DEFAULT_LOCAL_BASE / f"sch2visio_{stamp}"),
        help="Local directory for downloaded inputs and generated Visio output.",
    )
    parser.add_argument(
        "--output-vsdx",
        help="Final .vsdx path. Defaults to <local-dir>/schematic.vsdx.",
    )
    parser.add_argument(
        "--skill-timeout",
        type=int,
        default=300,
        help="Timeout in seconds for remote SKILL export.",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Stop after export, download, TSV-to-XLSX conversion, and dry validation.",
    )
    parser.add_argument(
        "--show-visio",
        action="store_true",
        help="Show the Visio window while drawing. Default is hidden.",
    )
    parser.add_argument(
        "extra_visio_args",
        nargs=argparse.REMAINDER,
        help="Extra arguments forwarded to cadence_to_visio_v2.py after --.",
    )
    return parser.parse_args()


def forwarded_args(raw_args: list[str]) -> list[str]:
    if raw_args and raw_args[0] == "--":
        return raw_args[1:]
    return raw_args


def find_bridge_root(raw: str | None) -> Path:
    candidates = []
    if raw:
        candidates.append(Path(raw))
    candidates.extend(
        [
            Path.cwd() / "virtuoso-bridge-lite",
            Path.cwd().parent / "virtuoso-bridge-lite",
        ]
    )
    for candidate in candidates:
        if (candidate / "src" / "virtuoso_bridge").exists():
            return candidate.resolve()
    raise FileNotFoundError(
        "Cannot find virtuoso-bridge-lite. Pass --bridge-root C:\\path\\to\\virtuoso-bridge-lite."
    )


def skill_quote(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def check_result(result, action: str) -> None:
    status = str(getattr(result, "status", "")).lower()
    errors = getattr(result, "errors", None) or []
    if "error" in status or errors:
        detail = "; ".join(str(item) for item in errors) or str(getattr(result, "output", ""))
        raise RuntimeError(f"{action} failed: {detail}")


def legacy_scp_download(remote_path: str, local_path: Path) -> None:
    host = os.environ.get("VB_REMOTE_HOST")
    user = os.environ.get("VB_REMOTE_USER")
    if not host or not user:
        raise RuntimeError("VB_REMOTE_HOST/VB_REMOTE_USER are not set for legacy scp fallback")

    local_path.parent.mkdir(parents=True, exist_ok=True)
    target = f"{user}@{host}:{remote_path}"
    command = [
        "scp",
        "-O",
        "-o",
        "BatchMode=yes",
        "-o",
        "StrictHostKeyChecking=no",
        target,
        str(local_path),
    ]
    subprocess.run(command, check=True)


def download_text_artifact(client, remote_path: str, local_path: Path) -> None:
    result = client.download_file(remote_path, local_path, timeout=120)
    try:
        check_result(result, f"download {local_path.name}")
        return
    except RuntimeError as exc:
        print(f"{exc}; retrying with legacy scp -O")
    legacy_scp_download(remote_path, local_path)


def run_remote_export(args: argparse.Namespace, bridge_root: Path) -> None:
    sys.path.insert(0, str(bridge_root / "src"))
    from virtuoso_bridge import VirtuosoClient  # noqa: PLC0415

    client = VirtuosoClient.from_env()

    load_expr = f'load("{skill_quote(args.remote_skill)}")'
    export_expr = f'sch2visio("{skill_quote(args.remote_dir)}")'

    print(f"Loading remote SKILL: {args.remote_skill}")
    check_result(client.execute_skill(load_expr, timeout=60), "load remote sch2visio.il")

    print(f"Exporting current schematic to remote dir: {args.remote_dir}")
    check_result(
        client.execute_skill(export_expr, timeout=args.skill_timeout),
        "remote sch2visio export",
    )

    local_dir = Path(args.local_dir).resolve()
    local_dir.mkdir(parents=True, exist_ok=True)
    downloads = {
        "netlist.cdl": local_dir / "netlist.cdl",
        "inst_info.txt": local_dir / "inst_info.txt",
        "wires.tsv": local_dir / "wires.tsv",
    }
    for remote_name, local_path in downloads.items():
        remote_path = f"{args.remote_dir}/{remote_name}"
        print(f"Downloading {remote_path} -> {local_path}")
        download_text_artifact(client, remote_path, local_path)

def run_cadence_to_visio(args: argparse.Namespace) -> Path:
    if not CADENCE_TO_VISIO_WRAPPER.exists():
        raise FileNotFoundError(f"Missing bundled Visio wrapper: {CADENCE_TO_VISIO_WRAPPER}")

    local_dir = Path(args.local_dir).resolve()

    common = [
        "--wires",
        str(local_dir / "wires.tsv"),
        "--netlist",
        str(local_dir / "netlist.cdl"),
        "--inst-info",
        str(local_dir / "inst_info.txt"),
        "--cwd",
        str(local_dir),
    ]

    validate_cmd = [sys.executable, str(CADENCE_TO_VISIO_WRAPPER), "validate", *common]
    print("Validating Visio inputs")
    subprocess.run(validate_cmd, check=True)

    output_vsdx = Path(args.output_vsdx).resolve() if args.output_vsdx else local_dir / "schematic.vsdx"
    if args.validate_only:
        print(f"Validation complete. Local artifacts: {local_dir}")
        return output_vsdx

    native_args = forwarded_args(args.extra_visio_args)
    if not args.show_visio and "--hidden" not in native_args:
        native_args.append("--hidden")

    visio_cmd = [sys.executable, str(CADENCE_TO_VISIO_WRAPPER), "visio", *common]
    if native_args:
        visio_cmd.extend(["--", *native_args])

    print("Drawing schematic in Visio")
    subprocess.run(visio_cmd, check=True)
    save_active_visio_document(output_vsdx)
    return output_vsdx


def save_active_visio_document(output_vsdx: Path) -> None:
    try:
        import win32com.client  # noqa: PLC0415
    except ImportError as exc:
        raise RuntimeError("pywin32 is required to save the generated Visio document") from exc

    output_vsdx.parent.mkdir(parents=True, exist_ok=True)
    try:
        visio = win32com.client.GetActiveObject("Visio.Application")
    except Exception:
        visio = win32com.client.Dispatch("Visio.Application")

    document = visio.ActiveDocument
    document.SaveAs(str(output_vsdx))
    print(f"Saved Visio drawing: {output_vsdx}")


def main() -> int:
    args = parse_args()
    bridge_root = find_bridge_root(args.bridge_root)
    run_remote_export(args, bridge_root)
    output_vsdx = run_cadence_to_visio(args)
    if not args.validate_only:
        print(f"Done: {output_vsdx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
