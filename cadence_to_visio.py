import re
import win32com.client
import math

# === 配置 ===
INPUT_FILE   = r"C:\Wangzz\20251009\inst_info.txt"
NETLIST_FILE = r"C:\Wangzz\20251009\netlist.txt"
STENCIL      = r"C:\Wangzz\20251009\cadence_to_visio\circuit.vss"
SCALE        = 1

# 元件尺寸
# 元件尺寸
W_NMOS, H_NMOS = 0.44, 0.59
W_PMOS, H_PMOS = 0.44, 0.59
W_RES,  H_RES  = 0.20, 0.59   # 新增电阻尺寸
W_UNKNOWN, H_UNKNOWN = 0.25, 0.25

# === 连线模式开关 ===
STRICT_MODE = False   # False = 全连线模式；True = 严格模式
USE_LINE    = True    # True = 使用 DrawLine；False = 使用 Dynamic Connector


# 不参与连线的网络与引脚
EXCLUDED_NETS = {"VDDA", "VSSA", "GNDA"}
EXCLUDED_PINS = {"B"}

# === 方向应用到 Visio 形状 ===
def apply_orientation(shape, orient):
    angle_map = {
        "R0": 0,
        "R90": math.pi/2,
        "R180": math.pi,
        "R270": 3*math.pi/2,
    }
    if orient in angle_map:
        shape.CellsU("Angle").ResultIU = angle_map[orient]
    elif orient == "MX":
        shape.CellsU("FlipY").FormulaU = "1"
    elif orient == "MY":
        shape.CellsU("FlipX").FormulaU = "1"
    elif orient == "MXR90":
        shape.CellsU("FlipY").FormulaU = "1"
        shape.CellsU("Angle").ResultIU = math.pi/2
    elif orient == "MYR90":
        shape.CellsU("FlipX").FormulaU = "1"
        shape.CellsU("Angle").ResultIU = math.pi/2

# === 解析 inst_info.txt ===
def parse_instances(filename):
    instances = {}
    with open(filename, "r") as f:
        content = f.read()
    blocks = content.strip().split("\n\n")
    for block in blocks:
        name_m   = re.search(r"Name:\s+(\S+)", block)
        xy_m     = re.search(r"XY:\s+\((-?\d+\.?\d*)\s+(-?\d+\.?\d*)\)", block)
        orient_m = re.search(r"Orient:\s+(\S+)", block)
        if not (name_m and xy_m and orient_m):
            continue
        name   = name_m.group(1)
        x      = float(xy_m.group(1)) * SCALE
        y      = float(xy_m.group(2)) * SCALE
        orient = orient_m.group(1)

        # === 类型识别 ===
        if name.upper().startswith("NM") or name.upper().startswith("M"):
            dev_type = "NMOS"
        elif name.upper().startswith("PM"):
            dev_type = "PMOS"
        elif name.upper().startswith("R"):
            dev_type = "RES"
        else:
            dev_type = "UNKNOWN"

        instances[name] = {
            "name": name,
            "type": dev_type,
            "xy": (x, y),
            "orient": orient
        }
    return instances

# === 解析 netlist.txt ===
def parse_netlist(filename):
    devices = []
    with open(filename, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("*") or line.startswith("."):
                continue
            tokens = line.split()
            if len(tokens) < 3:
                continue
            raw_name = tokens[0]
            model_idx = next((i for i, t in enumerate(tokens) if t.endswith("_ckt")), None)
            if model_idx is None or model_idx < 2:
                continue
            pins = tokens[1:model_idx]
            model = tokens[model_idx]
            name = raw_name[1:] if raw_name.startswith("X") else raw_name
            if name.upper().startswith("NM") or name.upper().startswith("M"):
                dev_type = "NMOS"
                pin_names = ["D", "G", "S", "B"]
            elif name.upper().startswith("PM"):
                dev_type = "PMOS"
                pin_names = ["D", "G", "S", "B"]
            elif name.upper().startswith("R"):
                dev_type = "RES"
                pin_names = ["R_up", "R_down"]
            else:
                dev_type = "UNKNOWN"
                pin_names = [f"P{i+1}" for i in range(len(pins))]
            pin_map = dict(zip(pin_names, pins))
            devices.append({
                "name": name,
                "type": dev_type,
                "pins": pin_map,
                "model": model
            })
    return devices

# === 引脚坐标（含旋转/镜像） ===
def get_pin_position(inst, pin, w, h):
    cx, cy = inst["xy"]
    orient = inst["orient"]

    if inst["type"] == "NMOS":
        local_map = {"D": ( w/2,  h/2), "G": (-w/2, 0), "S": ( w/2,-h/2), "B": ( w/2, 0)}
    elif inst["type"] == "PMOS":
        local_map = {"D": ( w/2,-h/2), "G": (-w/2, 0), "S": ( w/2, h/2), "B": ( w/2, 0)}
    elif inst["type"] == "RES":
        local_map = {"R_up": (0, h/2), "R_down": (0, -h/2)}
    else:
        return (cx, cy)

    if pin not in local_map:
        return (cx, cy)

    lx, ly = local_map[pin]
    # 旋转/镜像处理（和 MOS 一样）
    def rotate(x, y, angle):
        cos_a, sin_a = math.cos(angle), math.sin(angle)
        return (x*cos_a - y*sin_a, x*sin_a + y*cos_a)

    if orient == "R0": tx, ty = lx, ly
    elif orient == "R90": tx, ty = rotate(lx, ly, math.pi/2)
    elif orient == "R180": tx, ty = rotate(lx, ly, math.pi)
    elif orient == "R270": tx, ty = rotate(lx, ly, 3*math.pi/2)
    elif orient == "MX": tx, ty = lx, -ly
    elif orient == "MY": tx, ty = -lx, ly
    elif orient == "MXR90": tx, ty = rotate(lx, -ly, math.pi/2)
    elif orient == "MYR90": tx, ty = rotate(-lx, ly, math.pi/2)
    else: tx, ty = lx, ly

    return (cx + tx, cy + ty)


# === 包围盒与穿越检测 ===
def get_bbox(inst, w, h):
    cx, cy = inst["xy"]
    return (cx - w/2, cy - h/2, cx + w/2, cy + h/2)

def segment_crosses_bbox(p1, p2, bboxes, ignore_names):
    x1, y1 = p1; x2, y2 = p2
    horiz = abs(y1 - y2) < 1e-6
    vert  = abs(x1 - x2) < 1e-6
    if not (horiz or vert): return False
    xmin, xmax = min(x1, x2), max(x1, x2)
    ymin, ymax = min(y1, y2), max(y1, y2)
    for name,(bxmin,bymin,bxmax,bymax) in bboxes.items():
        if name in ignore_names: continue
        if horiz:
            if bymin <= y1 <= bymax and not (xmax < bxmin or xmin > bxmax):
                return True
        elif vert:
            if bxmin <= x1 <= bxmax and not (ymax < bymin or ymin > bymax):
                return True
    return False

def segment_hits_other_net_point(p1, p2, all_points, same_net_points):
    x1, y1 = p1; x2, y2 = p2
    horiz = abs(y1 - y2) < 1e-6
    vert  = abs(x1 - x2) < 1e-6
    if not (horiz or vert): return False
    xmin, xmax = min(x1, x2), max(x1, x2)
    ymin, ymax = min(y1, y2), max(y1, y2)
    same_set = {pt for pt in same_net_points}
    for (px, py) in all_points:
        if (px, py) in same_set:
            continue
        if horiz and abs(py - y1) < 1e-6 and xmin <= px <= xmax:
            return True
        if vert and abs(px - x1) < 1e-6 and ymin <= py <= ymax:
            return True
    return False


# === 放置器件并记录引脚位置 ===
def drop_with_label(page, master, inst, w, h, pin_positions, pin_list):
    cx, cy = inst["xy"]
    name = inst["name"]
    orient = inst["orient"]

    # 以中心放置形状
    shp = page.Drop(master, cx, cy)
    shp.Text = name
    shp.CellsU("Width").ResultIU  = w
    shp.CellsU("Height").ResultIU = h

    # 文本位置与尺寸
    shp.CellsU("TxtPinX").ResultIU   = shp.CellsU("Width").ResultIU + 0.20
    shp.CellsU("TxtPinY").ResultIU   = shp.CellsU("Height").ResultIU / 2.0
    shp.CellsU("TxtWidth").ResultIU  = 0.6
    shp.CellsU("TxtHeight").ResultIU = 0.2

    # 应用方向
    apply_orientation(shp, orient)

    # 记录引脚坐标
    for pin in pin_list:
        pin_positions[name + ":" + pin] = get_pin_position(inst, pin, w, h)

    return shp

# def draw_net_lines(page, netlist, pin_positions, instances, bboxes):
    net_to_points = {}
    all_points = []

    # 收集每个网络的点
    for dev in netlist:
        name = dev["name"]
        for pin, net in dev["pins"].items():
            if pin.upper() in EXCLUDED_PINS or net.upper() in EXCLUDED_NETS:
                continue
            key = name + ":" + pin
            if key in pin_positions:
                pt = pin_positions[key]
                net_to_points.setdefault(net, []).append((name, pin, pt))
                all_points.append(pt)

    # 遍历网络
    for net, pins in net_to_points.items():
        segments = []

        if STRICT_MODE:
            # === 严格模式：只画水平/垂直线，避免穿越 ===
            coords_same_net = [pt for _, _, pt in pins]
            for i in range(len(pins)):
                for j in range(i+1, len(pins)):
                    name1, pin1, p1 = pins[i]
                    name2, pin2, p2 = pins[j]

                    # 跳过同一器件的 D-S
                    if name1 == name2 and {pin1.upper(), pin2.upper()} == {"D", "S"}:
                        continue

                    horiz = abs(p1[1] - p2[1]) < 1e-6
                    vert  = abs(p1[0] - p2[0]) < 1e-6
                    if not (horiz or vert):
                        continue

                    if segment_crosses_bbox(p1, p2, bboxes, ignore_names={name1, name2}):
                        continue
                    if segment_hits_other_net_point(p1, p2, all_points, coords_same_net):
                        continue

                    x1, y1 = p1
                    x2, y2 = p2
                    if x1 > x2 or y1 > y2:
                        x1, y1, x2, y2 = x2, y2, x1, y1
                    segments.append(((x1, y1), (x2, y2)))
        else:
            # === 全连线模式：所有引脚两两相连 ===
            for i in range(len(pins)):
                for j in range(i+1, len(pins)):
                    _, _, p1 = pins[i]
                    _, _, p2 = pins[j]
                    segments.append((p1, p2))

        # 绘制
        for p1, p2 in segments:
            if USE_LINE:
                # 使用普通直线
                line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
                line.CellsU("LineWeight").FormulaU = "1.2 pt"
            else:
                # 使用 Dynamic Connector
                connector = page.Drop(page.Application.ConnectorToolDataObject, 0, 0)
                connector.CellsU("BeginX").ResultIU = p1[0]
                connector.CellsU("BeginY").ResultIU = p1[1]
                connector.CellsU("EndX").ResultIU   = p2[0]
                connector.CellsU("EndY").ResultIU   = p2[1]
                connector.CellsU("LineWeight").FormulaU = "1.2 pt"
                # 可选：强制直线风格
                connector.CellsU("RouteStyle").FormulaU = "16"



def build_mst(points, candidate_edges=None):
    """Kruskal MST，candidate_edges 可选：[(dist, i, j)]"""
    if candidate_edges is None:
        # 默认全连线候选
        candidate_edges = []
        for i, p1 in enumerate(points):
            for j, p2 in enumerate(points):
                if i < j:
                    dist = abs(p1[0]-p2[0]) + abs(p1[1]-p2[1])
                    candidate_edges.append((dist, i, j))
    candidate_edges.sort()

    parent = list(range(len(points)))
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    mst = []
    for dist, i, j in candidate_edges:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[ri] = rj
            mst.append((points[i], points[j]))
    return mst


FULL_CONNECT = False   # True = 全连线（可能冗余），False = 去冗余（MST）

def draw_net_lines(page, netlist, pin_positions, instances, bboxes):
    net_to_points = {}

    # 收集每个网络的点
    for dev in netlist:
        name = dev["name"]
        for pin, net in dev["pins"].items():
            if pin.upper() in EXCLUDED_PINS or net.upper() in EXCLUDED_NETS:
                continue
            key = name + ":" + pin
            if key in pin_positions:
                pt = pin_positions[key]
                net_to_points.setdefault(net, []).append((name, pin, pt))

    for net, pins in net_to_points.items():
        if len(pins) < 2:
            continue

        coords = [pt for _, _, pt in pins]

        if STRICT_MODE:
            # 严格模式：只考虑水平/垂直且不穿越器件的 MST
            candidate_edges = []
            for i, (n1, pin1, c1) in enumerate(pins):
                for j, (n2, pin2, c2) in enumerate(pins):
                    if i < j:
                        horiz = abs(c1[1]-c2[1]) < 1e-6
                        vert  = abs(c1[0]-c2[0]) < 1e-6
                        if not (horiz or vert):
                            continue
                        if segment_crosses_bbox(c1, c2, bboxes, ignore_names={n1, n2}):
                            continue
                        dist = abs(c1[0]-c2[0]) + abs(c1[1]-c2[1])
                        candidate_edges.append((dist, i, j))
            edges = build_mst(coords, candidate_edges)
        else:
            if FULL_CONNECT:
                # 全连线：所有引脚两两相连
                edges = []
                for i, p1 in enumerate(coords):
                    for j, p2 in enumerate(coords):
                        if i < j:
                            edges.append((p1, p2))
            else:
                # 去冗余：用 MST
                edges = build_mst(coords)

        # 绘制
        for p1, p2 in edges:
            horiz = abs(p1[1]-p2[1]) < 1e-6
            vert  = abs(p1[0]-p2[0]) < 1e-6
            is_manhattan = horiz or vert

            if USE_LINE:
                line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
                if is_manhattan:
                    line.CellsU("LineWeight").FormulaU = "1.2 pt"
                else:
                    line.CellsU("LineWeight").FormulaU = "1.2 pt"
                    line.CellsU("LinePattern").FormulaU = "2"  # 虚线

            else:
                connector = page.Drop(page.Application.ConnectorToolDataObject, 0, 0)
                connector.CellsU("BeginX").ResultIU = p1[0]
                connector.CellsU("BeginY").ResultIU = p1[1]
                connector.CellsU("EndX").ResultIU   = p2[0]
                connector.CellsU("EndY").ResultIU   = p2[1]
                if is_manhattan:
                    connector.CellsU("LineWeight").FormulaU = "1.2 pt"
                    connector.CellsU("RouteStyle").FormulaU = "64"  # 避让障碍物的直角折线
                else:
                    connector.CellsU("LineWeight").FormulaU = "0.6 pt"
                    connector.CellsU("LinePattern").FormulaU = "2"  # 虚线
                    connector.CellsU("ConLineRouteExt").FormulaU = "2"
                    connector.CellsU("RouteStyle").FormulaU = "32"  # 直线

    page.Layout()





# === 主程序 ===
def main():
    visio = win32com.client.Dispatch("Visio.Application")
    visio.Visible = True
    doc = visio.Documents.Add("")
    page = visio.ActivePage

    stencil = visio.Documents.OpenEx(STENCIL, 64)
    M_NMOS    = stencil.Masters("NMOS")
    M_PMOS    = stencil.Masters("PMOS")
    M_RES     = stencil.Masters("R")        # 确认模具里电阻的 Master 名称
    M_UNKNOWN = stencil.Masters("Unknown")
    M_LINE    = stencil.Masters("Line")

    instances = parse_instances(INPUT_FILE)
    netlist   = parse_netlist(NETLIST_FILE)
    pin_positions = {}
    bboxes = {}

    # 放置器件并记录 bbox
    for inst in instances.values():
        name = inst["name"]
        if inst["type"] == "NMOS":
            drop_with_label(page, M_NMOS, inst, W_NMOS, H_NMOS, pin_positions, ["D","G","S","B"])
            bboxes[name] = get_bbox(inst, W_NMOS, H_NMOS)
        elif inst["type"] == "PMOS":
            drop_with_label(page, M_PMOS, inst, W_PMOS, H_PMOS, pin_positions, ["D","G","S","B"])
            bboxes[name] = get_bbox(inst, W_PMOS, H_PMOS)
        elif inst["type"] == "RES":
            drop_with_label(page, M_RES, inst, W_RES, H_RES, pin_positions, ["R_up","R_down"])
            bboxes[name] = get_bbox(inst, W_RES, H_RES)
        else:
            drop_with_label(page, M_UNKNOWN, inst, W_UNKNOWN, H_UNKNOWN, pin_positions, [])
            bboxes[name] = get_bbox(inst, W_UNKNOWN, H_UNKNOWN)

    print("所有器件已放置完成。")

    # 自动连线
    mode_text = "严格模式" if STRICT_MODE else "全连线模式"
    line_text = "直线 (Line)" if USE_LINE else "动态连接线 (Connector)"
    print(f"开始自动连线：{mode_text} + {line_text}")

    draw_net_lines(page, netlist, pin_positions, instances, bboxes)

    print("所有器件与连线已完成。")

        # 触发页面重新布局，让 Connector 避让生效
    page.Layout()


    # === 交互式处理虚线 ===
    choice = input("是否要将剩余虚线改为实线？(Y/N): ").strip().lower()
    if choice == "y":
        for shape in page.Shapes:
            try:
                if shape.CellExistsU("LinePattern", 0):
                    pattern = shape.CellsU("LinePattern").ResultIU
                    if pattern == 2:  # 虚线
                        shape.CellsU("LinePattern").FormulaU = "1"   # 改为实线
                        shape.CellsU("LineWeight").FormulaU = "1.2 pt"
            except Exception:
                pass
        print("已将剩余虚线改为实线。")
    else:
        print("保留虚线，不做修改。")




if __name__ == "__main__":
    main()
