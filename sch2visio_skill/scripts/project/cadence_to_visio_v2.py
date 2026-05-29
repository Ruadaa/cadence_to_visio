"""Cadence/Virtuoso schematic to Visio V2.0.

主入口脚本：
- 保留 Virtuoso 导出的 wire 端点坐标；
- 可选择是否把线端附着到器件 pin / 共享连接点；
- 可选择是否使用 Visio 内置 Dynamic Connector；
- 默认使用更稳定的普通 1D 线段，避免 Visio 自动重路由改变版图形状。
"""

from __future__ import annotations

import argparse
import math
import os
import shutil
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Set, Tuple

import cadence_to_visio_core as core


Point = Tuple[float, float]
EndpointKey = Tuple[str, Point]

ANCHOR_SIZE = 0.01
PIN_DIRECT_GLUE_EPS = core.PIN_ADAPTER_EPS


@dataclass
class GlueTarget:
    """Visio Glue 目标：可以是隐藏连接点，也可以是器件连接点。"""

    shape: object
    conn_idx: Optional[int] = None


@dataclass
class PinGlueCandidate:
    """候选器件 pin，用于匹配 wire endpoint 是否能直接附着。"""

    inst_name: str
    pin: str
    net: str
    shape: object
    conn_idx: int
    virt_point: Point
    page_point: Point


@dataclass
class PinGlueInfo:
    """已经确认可直接 Glue 的器件 pin 匹配结果。"""

    target: GlueTarget
    inst_name: str
    pin: str
    net: str
    distance: float


def set_formula(shape, cell_name: str, formula: str) -> None:
    try:
        shape.CellsU(cell_name).FormulaU = formula
    except Exception:
        pass


def set_result(shape, cell_name: str, value: float) -> None:
    try:
        shape.CellsU(cell_name).ResultIU = value
    except Exception:
        pass


def endpoint_key(segment: core.WireSegment, point: Point) -> EndpointKey:
    """同一 net + 同一坐标代表同一个电气连接点。"""

    return (core.normalize_net(segment.net), core.point_key(point))


def make_endpoint_anchor(page, point: Point):
    """创建隐藏连接点。

    线段之间、交汇点 node、无法直接附着到器件的端点，都会 Glue 到这个
    共享 anchor。anchor 不显示，但 Visio 会把它当作可附着对象。
    """

    x, y = point
    half = ANCHOR_SIZE / 2
    anchor = page.DrawRectangle(x - half, y - half, x + half, y + half)
    anchor.Text = ""
    set_formula(anchor, "LinePattern", "0")
    set_formula(anchor, "FillPattern", "0")
    set_formula(anchor, "LockWidth", "1")
    set_formula(anchor, "LockHeight", "1")
    set_formula(anchor, "LockMoveX", "1")
    set_formula(anchor, "LockMoveY", "1")
    set_formula(anchor, "LockSelect", "1")
    return anchor


def ensure_anchor(
    page,
    anchors: Dict[EndpointKey, object],
    key: EndpointKey,
    point: Point,
):
    anchor = anchors.get(key)
    if anchor is None:
        anchor = make_endpoint_anchor(page, point)
        anchors[key] = anchor
    return anchor


def unique_wire_endpoints(wires: Sequence[core.WireSegment]) -> Dict[EndpointKey, Point]:
    endpoints: Dict[EndpointKey, Point] = {}
    for segment in wires:
        for point in (segment.p1, segment.p2):
            endpoints[endpoint_key(segment, point)] = point
    return endpoints


def collect_pin_candidates(
    page,
    devices: Sequence[core.NetDevice],
    instances: Dict[str, core.Instance],
    shapes: Dict[str, object],
    coord: core.CoordMap,
    excluded_pins: Set[str],
) -> Dict[str, List[PinGlueCandidate]]:
    """读取每个器件 master 的 Connections.Xn/Yn，建立 pin 候选表。"""

    candidates_by_net: Dict[str, List[PinGlueCandidate]] = {}
    for device in devices:
        inst = instances.get(device.name)
        shape = shapes.get(device.name)
        if inst is None or shape is None:
            continue

        pin_order = core.DEVICE_PIN_ORDER.get(inst.dev_type, [])
        for pin, net in device.pins.items():
            if core.pin_is_excluded(inst.dev_type, pin, excluded_pins) or pin not in pin_order:
                continue

            conn_idx = pin_order.index(pin) + 1
            page_point = core.connection_page_point(page, shape, conn_idx)
            if page_point is None:
                continue

            candidate = PinGlueCandidate(
                inst_name=device.name,
                pin=pin,
                net=net,
                shape=shape,
                conn_idx=conn_idx,
                virt_point=coord.inverse_point(page_point),
                page_point=page_point,
            )
            candidates_by_net.setdefault(core.normalize_net(net), []).append(candidate)

    return candidates_by_net


def build_pin_targets(
    page,
    devices: Sequence[core.NetDevice],
    instances: Dict[str, core.Instance],
    shapes: Dict[str, object],
    wires: Sequence[core.WireSegment],
    coord: core.CoordMap,
    threshold: float,
    excluded_pins: Set[str],
) -> Tuple[Dict[EndpointKey, PinGlueInfo], int]:
    """匹配 wire endpoint 与器件 pin。

    只有坐标真正重合的 endpoint 才直接 Glue 到器件 pin。近但不重合的点
    会被跳过，避免 Visio 把原始线段端点拉到器件 pin 导致斜线/变形。
    """

    candidates_by_net = collect_pin_candidates(page, devices, instances, shapes, coord, excluded_pins)
    targets: Dict[EndpointKey, PinGlueInfo] = {}
    skipped_noncoincident = 0

    for key, endpoint_point in unique_wire_endpoints(wires).items():
        net_key, _point = key
        candidates = candidates_by_net.get(net_key, [])
        if not candidates:
            continue

        best = min(candidates, key=lambda candidate: core.distance_sq(endpoint_point, candidate.virt_point))
        distance = math.sqrt(core.distance_sq(endpoint_point, best.virt_point))
        if distance > threshold:
            continue
        if distance > PIN_DIRECT_GLUE_EPS:
            skipped_noncoincident += 1
            continue

        targets[key] = PinGlueInfo(
            target=GlueTarget(best.shape, best.conn_idx),
            inst_name=best.inst_name,
            pin=best.pin,
            net=best.net,
            distance=distance,
        )

    return targets, skipped_noncoincident


def glue_endpoint_to_target(shape, endpoint: str, target: GlueTarget) -> bool:
    """把 1D 线段或连接线的 Begin/End 端点 Glue 到目标。"""

    x_cell = shape.CellsU(f"{endpoint}X")
    y_cell = shape.CellsU(f"{endpoint}Y")

    if target.conn_idx is not None:
        try:
            x_cell.GlueTo(target.shape.CellsU(f"Connections.X{target.conn_idx}"))
            y_cell.GlueTo(target.shape.CellsU(f"Connections.Y{target.conn_idx}"))
            return True
        except Exception:
            pass

    try:
        x_cell.GlueTo(target.shape.CellsU("PinX"))
        y_cell.GlueTo(target.shape.CellsU("PinY"))
        return True
    except Exception:
        return False


def glue_shape_center_to_target(shape, target: GlueTarget) -> bool:
    """把 node 圆点中心附着到共享连接点或器件 pin。"""

    if target.conn_idx is not None:
        try:
            shape.CellsU("PinX").GlueTo(target.shape.CellsU(f"Connections.X{target.conn_idx}"))
            shape.CellsU("PinY").GlueTo(target.shape.CellsU(f"Connections.Y{target.conn_idx}"))
            return True
        except Exception:
            pass

    try:
        shape.CellsU("PinX").GlueTo(target.shape.CellsU("PinX"))
        shape.CellsU("PinY").GlueTo(target.shape.CellsU("PinY"))
        return True
    except Exception:
        return False


def endpoint_target(
    page,
    anchors: Dict[EndpointKey, object],
    pin_targets: Dict[EndpointKey, PinGlueInfo],
    segment: core.WireSegment,
    virt_point: Point,
    page_point: Point,
    attach: bool,
) -> Optional[GlueTarget]:
    if not attach:
        return None

    key = endpoint_key(segment, virt_point)
    pin_target = pin_targets.get(key)
    if pin_target is not None:
        return pin_target.target

    anchor = ensure_anchor(page, anchors, key, page_point)
    return GlueTarget(anchor)


def format_wire_shape(shape, use_visio_connectors: bool, weight: str = core.WIRE_WEIGHT) -> None:
    shape.Text = ""
    set_formula(shape, "LineWeight", weight)
    set_formula(shape, "LineColor", core.WIRE_COLOR)
    set_formula(shape, "LinePattern", "1")
    set_formula(shape, "BeginArrow", "0")
    set_formula(shape, "EndArrow", "0")
    if use_visio_connectors:
        # Visio 内置连接线会参与路由；默认关闭该模式，避免自动改线。
        set_formula(shape, "RouteStyle", "64")
        set_formula(shape, "ConLineJumpCode", "0")
        set_formula(shape, "ConLineJumpStyle", "0")


def draw_wire_shape(
    visio,
    page,
    begin: Point,
    end: Point,
    use_visio_connectors: bool,
):
    if use_visio_connectors:
        mid_x = (begin[0] + end[0]) / 2
        mid_y = (begin[1] + end[1]) / 2
        shape = page.Drop(visio.ConnectorToolDataObject, mid_x, mid_y)
        set_result(shape, "BeginX", begin[0])
        set_result(shape, "BeginY", begin[1])
        set_result(shape, "EndX", end[0])
        set_result(shape, "EndY", end[1])
    else:
        shape = page.DrawLine(begin[0], begin[1], end[0], end[1])

    format_wire_shape(shape, use_visio_connectors)
    return shape


def draw_wires(
    visio,
    page,
    wires: Sequence[core.WireSegment],
    coord: core.CoordMap,
    attach: bool,
    use_visio_connectors: bool,
    pin_targets: Dict[EndpointKey, PinGlueInfo],
) -> Tuple[Dict[EndpointKey, object], int, int]:
    anchors: Dict[EndpointKey, object] = {}
    glue_failures = 0

    for segment in wires:
        p1 = coord.point(segment.p1)
        p2 = coord.point(segment.p2)
        shape = draw_wire_shape(visio, page, p1, p2, use_visio_connectors)

        begin_target = endpoint_target(page, anchors, pin_targets, segment, segment.p1, p1, attach)
        end_target = endpoint_target(page, anchors, pin_targets, segment, segment.p2, p2, attach)
        if begin_target is not None and not glue_endpoint_to_target(shape, "Begin", begin_target):
            glue_failures += 1
        if end_target is not None and not glue_endpoint_to_target(shape, "End", end_target):
            glue_failures += 1

        try:
            shape.SendToBack()
        except Exception:
            pass

    return anchors, len(wires), glue_failures


def t_junction_vertex_keys(wires: Sequence[core.WireSegment]) -> List[EndpointKey]:
    wire_lookup, endpoint_to_segments = core.build_wire_graph_by_net(wires)
    return sorted(
        (
            vertex_key
            for vertex_key in endpoint_to_segments
            if core.is_t_junction_vertex(vertex_key, wire_lookup, endpoint_to_segments)
        ),
        key=lambda item: (item[0], item[1][0], item[1][1]),
    )


def draw_nodes(
    page,
    master_index: Dict[str, object],
    wires: Sequence[core.WireSegment],
    coord: core.CoordMap,
    attach: bool,
    anchors: Dict[EndpointKey, object],
    pin_targets: Dict[EndpointKey, PinGlueInfo],
) -> Tuple[int, int]:
    glue_failures = 0
    count = 0

    for key in t_junction_vertex_keys(wires):
        net_key, point = key
        node = core.draw_node(page, master_index, coord, point)
        count += 1

        if not attach:
            continue

        pin_target = pin_targets.get(key)
        if pin_target is not None:
            target = pin_target.target
        else:
            anchor = ensure_anchor(page, anchors, (net_key, point), coord.point(point))
            target = GlueTarget(anchor)

        if not glue_shape_center_to_target(node, target):
            glue_failures += 1

    return count, glue_failures


def draw_labels(page, instances: Dict[str, core.Instance], shapes: Dict[str, object]) -> int:
    count = 0
    for name, inst in instances.items():
        shape = shapes.get(name)
        if shape is None:
            continue
        core.draw_device_label(page, inst, shape)
        count += 1
    return count


def bring_to_front(shapes: Dict[str, object]) -> None:
    for shape in shapes.values():
        try:
            shape.BringToFront()
        except Exception:
            pass


def draw_visio(
    instances: Dict[str, core.Instance],
    devices: Sequence[core.NetDevice],
    wires: Sequence[core.WireSegment],
    args: argparse.Namespace,
    coord: core.CoordMap,
    excluded_pins: Set[str],
    placement_offsets: Sequence[core.PlacementOffsetRule],
) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("需要安装 pywin32：pip install pywin32") from exc

    try:
        visio = win32com.client.GetActiveObject("Visio.Application")
    except Exception:
        visio = win32com.client.Dispatch("Visio.Application")
    visio.Visible = not args.hidden
    visio.Documents.Add("")
    page = visio.ActivePage

    min_x, min_y, max_x, max_y = coord.bbox(coord.bounds)
    page.PageSheet.CellsU("PageWidth").ResultIU = max(1.0, max_x - min_x + core.PAGE_MARGIN)
    page.PageSheet.CellsU("PageHeight").ResultIU = max(1.0, max_y - min_y + core.PAGE_MARGIN)

    stencil = None
    if args.stencil and os.path.exists(args.stencil):
        stencil_path = os.path.abspath(args.stencil).lower()
        stencil_open_path = os.path.abspath(args.stencil)
        runtime_dir = os.path.dirname(os.path.abspath(args.wires))
        os.makedirs(runtime_dir, exist_ok=True)
        stencil_copy_path = os.path.join(runtime_dir, "_runtime_circuit.vss")
        shutil.copyfile(stencil_open_path, stencil_copy_path)
        stencil_open_path = stencil_copy_path
        for document in visio.Documents:
            try:
                if os.path.abspath(document.FullName).lower() == stencil_path:
                    stencil = document
                    break
            except Exception:
                continue
        if stencil is None:
            try:
                stencil = visio.Documents.OpenEx(stencil_open_path, 64)
            except Exception:
                for document in visio.Documents:
                    try:
                        full_name = getattr(document, "FullName", "") or ""
                        name = getattr(document, "Name", "") or ""
                        if (
                            full_name and os.path.abspath(full_name).lower() == stencil_path
                        ) or name.lower() == os.path.basename(stencil_path):
                            stencil = document
                            break
                    except Exception:
                        continue
                if stencil is None:
                    raise
    master_index = core.build_master_index(stencil)

    # 先放器件是为了读取 Connections.Xn/Yn；最后再 BringToFront 保证视觉上器件在最上层。
    shapes: Dict[str, object] = {}
    for inst in instances.values():
        shapes[inst.name] = core.draw_instance(
            page,
            master_index,
            inst,
            coord,
            args.symbol_fit,
            placement_offsets,
            draw_label=False,
        )

    if args.wire_adjust == "snap-endpoints":
        adjustments = core.snap_dangling_wire_endpoints(
            page,
            wires,
            devices,
            instances,
            shapes,
            coord,
            args.pin_snap_threshold,
            excluded_pins,
            placement_offsets,
        )
        core.write_wires_tsv(args.adjusted_wires_output, wires)
        core.write_adjustments_tsv(args.adjustment_report, adjustments)
        print(f"端点吸附：{len(adjustments)} 个 endpoint 被微调")

    if args.attach:
        pin_targets, skipped_noncoincident = build_pin_targets(
            page,
            devices,
            instances,
            shapes,
            wires,
            coord,
            args.pin_snap_threshold,
            excluded_pins,
        )
    else:
        pin_targets = {}
        skipped_noncoincident = 0

    anchors, wire_count, wire_glue_failures = draw_wires(
        visio,
        page,
        wires,
        coord,
        attach=args.attach,
        use_visio_connectors=args.visio_connectors,
        pin_targets=pin_targets,
    )

    if args.draw_nodes:
        node_count, node_glue_failures = draw_nodes(
            page,
            master_index,
            wires,
            coord,
            attach=args.attach,
            anchors=anchors,
            pin_targets=pin_targets,
        )
    else:
        node_count = 0
        node_glue_failures = 0

    bring_to_front(shapes)
    label_count = draw_labels(page, instances, shapes)
    glue_failures = wire_glue_failures + node_glue_failures

    print(f"绘制器件：{len(instances)} 个")
    print(f"绘制线段：{wire_count} 条")
    print(f"使用 Visio 内置连接线：{'是' if args.visio_connectors else '否'}")
    print(f"启用附着：{'是' if args.attach else '否'}")
    print(f"直接附着到器件 pin 的 endpoint：{len(pin_targets)} 个")
    print(f"为保留原始坐标而跳过的近邻 pin：{skipped_noncoincident} 个")
    print(f"共享隐藏连接点：{len(anchors)} 个")
    print(f"Glue 失败：{glue_failures} 个")
    print(f"Node：{node_count} 个；文本标签：{label_count} 个")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Cadence/Virtuoso 原理图转 Visio V2.0")
    parser.add_argument("--wires", default=core.DEFAULT_WIRES_FILE, help="wire 坐标文件，支持 wires.tsv 或 wires.xlsx")
    parser.add_argument("--netlist", default=core.DEFAULT_NETLIST_FILE, help="CDL 网表文件")
    parser.add_argument("--inst-info", default=core.DEFAULT_INST_INFO_FILE, help="器件坐标/方向信息")
    parser.add_argument("--stencil", default=core.DEFAULT_STENCIL_FILE, help="Visio stencil 文件")
    parser.add_argument("--placement-offsets", default=core.DEFAULT_PLACEMENT_OFFSETS_FILE, help="可选器件偏移表")
    parser.add_argument("--scale", type=float, default=1.0, help="全局缩放")
    parser.add_argument("--symbol-fit", choices=core.SYMBOL_FIT_MODES, default="native", help="器件 master 缩放方式")
    parser.add_argument("--wire-adjust", choices=("none", "snap-endpoints"), default="none", help="是否微调悬空端点")
    parser.add_argument("--pin-snap-threshold", type=float, default=0.8, help="pin/endpoint 匹配距离阈值")
    parser.add_argument("--exclude-pins", default="", help="排除附着/吸附的 pin，例如 B 或 MOS:B")
    parser.add_argument("--skip-nets", default="", help="完全跳过指定 net，例如 vdd,vss")
    parser.add_argument("--skip-mos-body-nets", action="store_true", help="跳过所有连接 MOS B 的 net")
    parser.add_argument(
        "--draw-mos-b-wires",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="是否绘制 MOS B 端分支线；默认不绘制",
    )
    parser.add_argument(
        "--draw-nodes",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="是否绘制 T 形交汇点 node",
    )
    parser.add_argument(
        "--attach",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="全局开关：是否把线段端点/node 附着到器件 pin 或共享连接点",
    )
    parser.add_argument(
        "--visio-connectors",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="全局开关：是否使用 Visio 内置 Dynamic Connector；默认否，避免自动重路由",
    )
    parser.add_argument("--adjusted-wires-output", default="wires_adjusted.tsv", help="端点微调后的 wire 输出")
    parser.add_argument("--adjustment-report", default="wire_adjustments.tsv", help="端点微调报告")
    parser.add_argument("--preserve-absolute", action="store_true", help="保留原始绝对坐标")
    parser.add_argument("--flip-y", action="store_true", help="翻转 Y 轴")
    parser.add_argument("--dry-run", action="store_true", help="只检查输入，不打开 Visio")
    parser.add_argument("--hidden", action="store_true", help="隐藏 Visio 窗口")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    instances = core.parse_instances(args.inst_info)
    devices = core.parse_netlist(args.netlist, instances)
    wires = core.read_wires(args.wires)
    placement_offsets = core.parse_placement_offsets(args.placement_offsets)

    for message in core.validate_inputs(instances, devices, wires):
        print(message)

    skip_nets = core.parse_net_set(args.skip_nets)
    if args.skip_mos_body_nets:
        body_nets = core.mos_body_nets(devices)
        skip_nets.update(body_nets)
        print(f"MOS B net：{', '.join(sorted(body_nets)) if body_nets else '(none)'}")

    wires, skipped_wire_count = core.filter_wires_by_net(wires, skip_nets)
    if skip_nets:
        print(f"跳过 net：{', '.join(sorted(skip_nets))}")
        print(f"跳过线段：{skipped_wire_count} 条；剩余 {len(wires)} 条")

    wires, skipped_body_wire_count, matched_body_pin_count = core.filter_mos_body_wire_branches(
        wires,
        devices,
        instances,
        draw_mos_body_wires=args.draw_mos_b_wires,
        pin_vertex_threshold=args.pin_snap_threshold,
        placement_offsets=placement_offsets,
    )
    if not args.draw_mos_b_wires:
        print(
            "MOS B 分支过滤："
            f"匹配 {matched_body_pin_count} 个 B pin，"
            f"跳过 {skipped_body_wire_count} 条线段，"
            f"剩余 {len(wires)} 条"
        )

    bounds = core.combined_bounds(instances.values(), wires, placement_offsets)
    coord = core.CoordMap(
        bounds=bounds,
        preserve_absolute=args.preserve_absolute,
        flip_y=args.flip_y,
        scale=args.scale,
    )

    print(f"Virtuoso bounds: x={bounds[0]}..{bounds[2]}, y={bounds[1]}..{bounds[3]}")
    if placement_offsets:
        print(f"读取器件偏移规则：{len(placement_offsets)} 条")
    if args.dry_run:
        print("Dry run 完成；未打开 Visio")
        return

    excluded_pins = {pin.strip().upper() for pin in args.exclude_pins.split(",") if pin.strip()}
    if not args.draw_mos_b_wires:
        excluded_pins.add("MOS:B")

    draw_visio(instances, devices, wires, args, coord, excluded_pins, placement_offsets)


if __name__ == "__main__":
    main()
