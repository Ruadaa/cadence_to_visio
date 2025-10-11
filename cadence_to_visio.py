import re
import win32com.client
import math

# === 配置 ===
INPUT_FILE   = r"C:\...\inst_info.txt"
NETLIST_FILE = r"C:\...\netlist.txt"
STENCIL      = r"C:\...\circuit.vss"
SCALE        = 1.5

# 元件尺寸
# 元件尺寸
W_NMOS, H_NMOS = 0.44, 0.59
W_PMOS, H_PMOS = 0.44, 0.59
W_RES,  H_RES  = 0.20, 0.59   # 新增电阻尺寸
W_UNKNOWN, H_UNKNOWN = 0.25, 0.25


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

# === 自动绘制连线（横/竖线；排除规则；跳过 D-S；去冗余；避免穿越） ===
def draw_net_lines(page, netlist, pin_positions, M_LINE, instances, bboxes):
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

                # 穿越器件？
                if segment_crosses_bbox(p1, p2, bboxes, ignore_names={name1, name2}):
                    continue

                # 穿过其他网络端点？
                if segment_hits_other_net_point(p1, p2, all_points, coords_same_net):
                    continue

                # 通过所有过滤
                x1, y1 = p1
                x2, y2 = p2
                if x1 > x2 or y1 > y2:
                    x1, y1, x2, y2 = x2, y2, x1, y1
                segments.append(((x1, y1), (x2, y2)))

        # 去除被包含的小线段
        filtered = []
        for i, (a1, a2) in enumerate(segments):
            keep = True
            for j, (b1, b2) in enumerate(segments):
                if i == j:
                    continue
                if a1[1] == a2[1] == b1[1] == b2[1]:  # 水平
                    if b1[0] <= a1[0] and a2[0] <= b2[0]:
                        keep = False
                        break
                elif a1[0] == a2[0] == b1[0] == b2[0]:  # 垂直
                    if b1[1] <= a1[1] and a2[1] <= b2[1]:
                        keep = False
                        break
            if keep:
                filtered.append((a1, a2))

        # 绘制
        for p1, p2 in filtered:
            line = page.Drop(M_LINE, 0, 0)
            line.CellsU("BeginX").ResultIU = p1[0]
            line.CellsU("BeginY").ResultIU = p1[1]
            line.CellsU("EndX").ResultIU   = p2[0]
            line.CellsU("EndY").ResultIU   = p2[1]


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
    draw_net_lines(page, netlist, pin_positions, M_LINE, instances, bboxes)

    print("所有器件与连线已完成。")

if __name__ == "__main__":
    main()
