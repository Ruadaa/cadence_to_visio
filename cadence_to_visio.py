import re
import win32com.client
import math

# === 配置 ===
INPUT_FILE   = r"C:\inst_info.txt"
NETLIST_FILE = r"C:\netlist.txt"
STENCIL      = r"C:\circuit.vss"
SCALE        = 1.5

# 元件尺寸
# 元件尺寸
W_NMOS, H_NMOS = 0.44, 0.59
W_PMOS, H_PMOS = 0.44, 0.59
W_RES,  H_RES  = 0.20, 0.59   # 新增电阻尺寸
W_UNKNOWN, H_UNKNOWN = 0.25, 0.25

# === 连线模式开关 ===
STRICT_MODE = True   # False = 全连线模式（推荐）；True = 严格模式(只连接横竖线)
USE_LINE    = True    # True = 使用 DrawLine；False = 使用 Dynamic Connector（推荐）


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

    # if inst["type"] == "NMOS":
    #     local_map = {"D": ( w/2,  h/2), "G": (-w/2, 0), "S": ( w/2,-h/2), "B": ( w/2, 0)}
    # elif inst["type"] == "PMOS":
    #     local_map = {"D": ( w/2,-h/2), "G": (-w/2, 0), "S": ( w/2, h/2), "B": ( w/2, 0)}
    # elif inst["type"] == "RES":
    #     local_map = {"R_up": (0, h/2), "R_down": (0, -h/2)}
    # else:
    #     return (cx, cy)

    if inst["type"] == "NMOS":
        # NMOS 精确端点
        local_map = {
            "D": (w * 1.0 - w/2,   h * 1.0    - h/2),   # Drain
            "G": (w * 0.0 - w/2,   h * 0.5017 - h/2),   # Gate
            "S": (w * 1.0 - w/2,   h * 0.0    - h/2),   # Source
            "B": (w * 0.5 - w/2,   h * 0.5    - h/2),   # Body（模具未定义，可放中心）
        }
    elif inst["type"] == "PMOS":
        # PMOS 精确端点（和 NMOS 对称）
        local_map = {
            "D": (w * 1.0 - w/2,   h * 0.0    - h/2),
            "G": (w * 0.0 - w/2,   h * 0.5017 - h/2),
            "S": (w * 1.0 - w/2,   h * 1.0    - h/2),
            "B": (w * 0.5 - w/2,   h * 0.5    - h/2),
        }
    elif inst["type"] == "RES":
        # 电阻精确端点
        local_map = {
            "R_up":   (w * 0.5 - w/2, h * 0.9867 - h/2),
            "R_down": (w * 0.5 - w/2, h * 0.0133 - h/2),
        }
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

# def draw_net_lines(page, netlist, pin_positions, instances, bboxes):
#     net_to_points = {}

#     # 收集每个网络的点
#     for dev in netlist:
#         name = dev["name"]
#         for pin, net in dev["pins"].items():
#             if pin.upper() in EXCLUDED_PINS or net.upper() in EXCLUDED_NETS:
#                 continue
#             key = name + ":" + pin
#             if key in pin_positions:
#                 pt = pin_positions[key]
#                 net_to_points.setdefault(net, []).append((name, pin, pt))

#     for net, pins in net_to_points.items():
#         if len(pins) < 2:
#             continue

#         coords = [pt for _, _, pt in pins]

#         if STRICT_MODE:
#             # 严格模式：只考虑水平/垂直且不穿越器件的 MST
#             candidate_edges = []
#             for i, (n1, pin1, c1) in enumerate(pins):
#                 for j, (n2, pin2, c2) in enumerate(pins):
#                     if i < j:
#                         horiz = abs(c1[1]-c2[1]) < 1e-6
#                         vert  = abs(c1[0]-c2[0]) < 1e-6
#                         if not (horiz or vert):
#                             continue
#                         if segment_crosses_bbox(c1, c2, bboxes, ignore_names={n1, n2}):
#                             continue
#                         dist = abs(c1[0]-c2[0]) + abs(c1[1]-c2[1])
#                         candidate_edges.append((dist, i, j))
#             edges = build_mst(coords, candidate_edges)
#         else:
#             if FULL_CONNECT:
#                 # 全连线：所有引脚两两相连
#                 edges = []
#                 for i, p1 in enumerate(coords):
#                     for j, p2 in enumerate(coords):
#                         if i < j:
#                             edges.append((p1, p2))
#             else:
#                 # 去冗余：用 MST
#                 edges = build_mst(coords)

#         # 绘制
#         for p1, p2 in edges:
#             horiz = abs(p1[1]-p2[1]) < 1e-6
#             vert  = abs(p1[0]-p2[0]) < 1e-6
#             is_manhattan = horiz or vert

#             if USE_LINE:
#                 line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
#                 if is_manhattan:
#                     line.CellsU("LineWeight").FormulaU = "1.2 pt"
#                 else:
#                     line.CellsU("LineWeight").FormulaU = "1.2 pt"
#                     line.CellsU("LinePattern").FormulaU = "2"  # 虚线

#             else:
#                 connector = page.Drop(page.Application.ConnectorToolDataObject, 0, 0)
#                 connector.CellsU("BeginX").ResultIU = p1[0]
#                 connector.CellsU("BeginY").ResultIU = p1[1]
#                 connector.CellsU("EndX").ResultIU   = p2[0]
#                 connector.CellsU("EndY").ResultIU   = p2[1]
#                 if is_manhattan:
#                     connector.CellsU("LineWeight").FormulaU = "1.2 pt"
#                     connector.CellsU("RouteStyle").FormulaU = "64"  # 避让障碍物的直角折线
#                 else:
#                     connector.CellsU("LineWeight").FormulaU = "0.6 pt"
#                     connector.CellsU("LinePattern").FormulaU = "2"  # 虚线
#                     connector.CellsU("ConLineRouteExt").FormulaU = "2"
#                     connector.CellsU("RouteStyle").FormulaU = "32"  # 直线

#     page.Layout()
# # 引脚到连接点的映射表
PIN_CONN_INDEX = {
    "NMOS": {"D": 1, "G": 2, "S": 3},
    "PMOS": {"D": 1, "G": 2, "S": 3},
    "R":    {"R_up": 1, "R_down": 2},
}

def glue_line_end(line, cell_name, dev_name, pin_name, instances):
    """把线的端点 Glue 到器件的连接点（同时 Glue X 和 Y）"""
    if dev_name not in instances:
        return
    shape = instances[dev_name]  # 这里的 instances 必须是 name -> Visio shape 的映射

    # 用 Master 名称选择映射，避免用实例名前缀误判
    master_name = shape.Master.NameU.upper()
    pin_map = PIN_CONN_INDEX.get(master_name, None)
    if not pin_map:
        # 兼容旧逻辑：用实例名前缀匹配（不推荐，但保留）
        for key in PIN_CONN_INDEX:
            if master_name.startswith(key):
                pin_map = PIN_CONN_INDEX[key]
                break
    if not pin_map:
        return

    idx = pin_map.get(pin_name)
    if not idx:
        return

    try:
        conn_x = shape.CellsU(f"Connections.X{idx}")
        conn_y = shape.CellsU(f"Connections.Y{idx}")
        if cell_name == "BeginX":
            line.CellsU("BeginX").GlueTo(conn_x)
            line.CellsU("BeginY").GlueTo(conn_y)
        else:
            line.CellsU("EndX").GlueTo(conn_x)
            line.CellsU("EndY").GlueTo(conn_y)
    except Exception as e:
        print(f"[Glue] {dev_name}:{pin_name} 失败: {e}")

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
                    line.CellsU("LineWeight").FormulaU = "0.6 pt"
                    line.CellsU("LinePattern").FormulaU = "2"

                # Glue
                dev1, pin1, _ = next((x for x in pins if x[2] == p1), (None, None, None))
                dev2, pin2, _ = next((x for x in pins if x[2] == p2), (None, None, None))
                if dev1 and pin1:
                    glue_line_end(line, "BeginX", dev1, pin1, instances)
                if dev2 and pin2:
                    glue_line_end(line, "EndX", dev2, pin2, instances)

            else:
                connector = page.Drop(page.Application.ConnectorToolDataObject, 0, 0)
                connector.CellsU("BeginX").ResultIU = p1[0]
                connector.CellsU("BeginY").ResultIU = p1[1]
                connector.CellsU("EndX").ResultIU   = p2[0]
                connector.CellsU("EndY").ResultIU   = p2[1]
                if is_manhattan:
                    connector.CellsU("LineWeight").FormulaU = "1.2 pt"
                    # connector.CellsU("RouteStyle").FormulaU = "64"
                else:
                    connector.CellsU("LineWeight").FormulaU = "1.2 pt"
                    connector.CellsU("LinePattern").FormulaU = "2"
                    # connector.CellsU("RouteStyle").FormulaU = "64"

                # Glue
                dev1, pin1, _ = next((x for x in pins if x[2] == p1), (None, None, None))
                dev2, pin2, _ = next((x for x in pins if x[2] == p2), (None, None, None))
                if dev1 and pin1:
                    glue_line_end(connector, "BeginX", dev1, pin1, instances)
                if dev2 and pin2:
                    glue_line_end(connector, "EndX", dev2, pin2, instances)


    # page.Layout()






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
    # 新增：收集已放置的 Visio shape
    shapes_map = {}


    # 放置器件并记录 bbox
    for inst in instances.values():
        name = inst["name"]
        if inst["type"] == "NMOS":
            shp = drop_with_label(page, M_NMOS, inst, W_NMOS, H_NMOS, pin_positions, ["D","G","S","B"])
            shapes_map[name] = shp
            bboxes[name] = get_bbox(inst, W_NMOS, H_NMOS)
        elif inst["type"] == "PMOS":
            shp = drop_with_label(page, M_PMOS, inst, W_PMOS, H_PMOS, pin_positions, ["D","G","S","B"])
            shapes_map[name] = shp
            bboxes[name] = get_bbox(inst, W_PMOS, H_PMOS)
        elif inst["type"] == "RES":
            shp = drop_with_label(page, M_RES, inst, W_RES, H_RES, pin_positions, ["R_up","R_down"])
            shapes_map[name] = shp
            bboxes[name] = get_bbox(inst, W_RES, H_RES)
        else:
            shp = drop_with_label(page, M_UNKNOWN, inst, W_UNKNOWN, H_UNKNOWN, pin_positions, [])
            shapes_map[name] = shp
            bboxes[name] = get_bbox(inst, W_UNKNOWN, H_UNKNOWN)


    print("所有器件已放置完成。")

    # 自动连线
    mode_text = "严格模式" if STRICT_MODE else "全连线模式"
    line_text = "直线 (Line)" if USE_LINE else "动态连接线 (Connector)"
    print(f"开始自动连线：{mode_text} + {line_text}")

    draw_net_lines(page, netlist, pin_positions, shapes_map, bboxes)


    print("所有器件与连线已完成。")

        # 触发页面重新布局，让 Connector 避让生效
    # page.Layout()


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
