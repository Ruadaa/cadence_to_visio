import re
import win32com.client
import math

# === 配置 ===
INPUT_FILE   = r"C:\inst_info1.txt"
NETLIST_FILE = r"C:\netlist1.txt"
STENCIL      = r"C:\circuit.vss"
SCALE        = 1.5

# 元件尺寸
W_NMOS, H_NMOS = 0.44, 0.59
W_PMOS, H_PMOS = 0.44, 0.59
W_UNKNOWN, H_UNKNOWN = 0.25, 0.25

# 不参与连线的网络与引脚
EXCLUDED_NETS = {"VDDA", "GNDA"}
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
        if name.upper().startswith("NM"):
            dev_type = "NMOS"
        elif name.upper().startswith("PM"):
            dev_type = "PMOS"
        else:
            dev_type = "UNKNOWN"
        instances[name] = {
            "name": name,
            "type": dev_type,
            "xy": (x, y),      # 作为器件中心点
            "orient": orient
        }
    return instances

# === 解析 netlist.txt（去掉器件名前缀 X） ===
def parse_netlist(filename):
    devices = []
    with open(filename, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("*") or line.startswith("."):
                continue

            # 分割整行
            tokens = line.split()
            if len(tokens) < 3:
                continue

            raw_name = tokens[0]
            model_idx = next((i for i, t in enumerate(tokens) if t.endswith("_ckt")), None)
            if model_idx is None or model_idx < 2:
                continue

            pins = tokens[1:model_idx]
            model = tokens[model_idx]

            # 去掉前缀 X
            name = raw_name[1:] if raw_name.startswith("X") else raw_name

            # 判断器件类型
            if name.upper().startswith("NM"):
                dev_type = "NMOS"
                pin_names = ["D", "G", "S", "B"]
            elif name.upper().startswith("PM"):
                dev_type = "PMOS"
                pin_names = ["D", "G", "S", "B"]
            elif name.upper().startswith("R"):
                dev_type = "RES"
                pin_names = ["1", "2"]
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

# === 方向/镜像下的引脚坐标变换（以器件中心为基准） ===
def get_pin_position(inst, pin, w, h):
    cx, cy = inst["xy"]   # 器件中心坐标（注意：drop 时用中心对齐）
    orient = inst["orient"]

    # R0 情况下的局部坐标（以中心为基准）
    if inst["type"] == "NMOS":
        local_map = {
            "D": ( w/2,  h/2),  # 右上角
            "G": (-w/2,  0   ), # 左中
            "S": ( w/2, -h/2),  # 右下角
            "B": ( w/2,  0   ), # 右中
        }
    elif inst["type"] == "PMOS":
        local_map = {
            "D": ( w/2, -h/2),  # 右下角
            "G": (-w/2,  0   ), # 左中
            "S": ( w/2,  h/2),  # 右上角
            "B": ( w/2,  0   ), # 右中
        }
    else:
        return (cx, cy)

    if pin not in local_map:
        return (cx, cy)

    lx, ly = local_map[pin]

    # 几何变换
    def rotate(x, y, angle):
        cos_a, sin_a = math.cos(angle), math.sin(angle)
        return (x*cos_a - y*sin_a, x*sin_a + y*cos_a)

    if orient == "R0":
        tx, ty = lx, ly
    elif orient == "R90":
        tx, ty = rotate(lx, ly, math.pi/2)
    elif orient == "R180":
        tx, ty = rotate(lx, ly, math.pi)
    elif orient == "R270":
        tx, ty = rotate(lx, ly, 3*math.pi/2)
    elif orient == "MX":   # Y 轴翻转
        tx, ty = lx, -ly
    elif orient == "MY":   # X 轴翻转
        tx, ty = -lx, ly
    elif orient == "MXR90":
        tx, ty = rotate(lx, -ly, math.pi/2)
    elif orient == "MYR90":
        tx, ty = rotate(-lx, ly, math.pi/2)
    else:
        tx, ty = lx, ly

    # 转换到全局坐标（中心平移）
    return (cx + tx, cy + ty)

# === 放置器件并记录引脚位置 ===
def drop_with_label(page, master, inst, w, h, pin_positions):
    cx, cy = inst["xy"]
    name = inst["name"]
    orient = inst["orient"]

    # 以中心放置形状（PinX/PinY 在中心）
    shp = page.Drop(master, cx, cy)
    shp.Text = name
    shp.CellsU("Width").ResultIU  = w
    shp.CellsU("Height").ResultIU = h

    # 文本位置与尺寸
    shp.CellsU("TxtPinX").ResultIU   = shp.CellsU("Width").ResultIU + 0.20
    shp.CellsU("TxtPinY").ResultIU   = shp.CellsU("Height").ResultIU / 2.0
    shp.CellsU("TxtWidth").ResultIU  = 0.6
    shp.CellsU("TxtHeight").ResultIU = 0.2

    # 应用形状方向
    apply_orientation(shp, orient)

    # 记录四个引脚的坐标（经过方向/镜像变换）
    for pin in ["D", "G", "S", "B"]:
        pin_positions[name + ":" + pin] = get_pin_position(inst, pin, w, h)

    return shp

# === 自动绘制连线（仅横线/竖线；排除 B/VSSA/VSSD） ===
def draw_net_lines(page, netlist, pin_positions, M_LINE):
    net_to_segments = {}

    # 构建每个 net 的线段列表（仅横/竖线）
    for dev in netlist:
        name = dev["name"]
        for pin, net in dev["pins"].items():
            if pin.upper() in EXCLUDED_PINS or net.upper() in EXCLUDED_NETS:
                continue
            key = name + ":" + pin
            if key in pin_positions:
                net_to_segments.setdefault(net, []).append((name, pin, pin_positions[key]))

    for net, pins in net_to_segments.items():
        segments = []
        for i in range(len(pins)):
            for j in range(i+1, len(pins)):
                name1, pin1, p1 = pins[i]
                name2, pin2, p2 = pins[j]

                # ✅ 跳过同一器件的 D-S 连线
                if name1 == name2 and {pin1.upper(), pin2.upper()} == {"D", "S"}:
                    continue

                if abs(p1[0] - p2[0]) < 1e-6 or abs(p1[1] - p2[1]) < 1e-6:
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
                # 同一方向
                if a1[1] == a2[1] == b1[1] == b2[1]:  # 水平线
                    if b1[0] <= a1[0] and a2[0] <= b2[0]:
                        keep = False
                        break
                elif a1[0] == a2[0] == b1[0] == b2[0]:  # 竖直线
                    if b1[1] <= a1[1] and a2[1] <= b2[1]:
                        keep = False
                        break
            if keep:
                filtered.append((a1, a2))

        # 画线
        for p1, p2 in filtered:
            line = page.Drop(M_LINE, 0, 0)
            line.CellsU("BeginX").ResultIU = p1[0]
            line.CellsU("BeginY").ResultIU = p1[1]
            line.CellsU("EndX").ResultIU   = p2[0]
            line.CellsU("EndY").ResultIU   = p2[1]
            # line.Text = net

# === 主程序 ===
def main():
    visio = win32com.client.Dispatch("Visio.Application")
    visio.Visible = True
    doc = visio.Documents.Add("")
    page = visio.ActivePage

    stencil = visio.Documents.OpenEx(STENCIL, 64)
    M_NMOS    = stencil.Masters("NMOS")
    M_PMOS    = stencil.Masters("PMOS")
    M_UNKNOWN = stencil.Masters("Unknown")
    M_LINE    = stencil.Masters("Line")  # 确认该 master 为动态连接器

    # 解析实例与网表
    instances = parse_instances(INPUT_FILE)
    netlist   = parse_netlist(NETLIST_FILE)
    pin_positions = {}

    # 放置器件并记录引脚位置（按中心放置）
    for inst in instances.values():
        if inst["type"] == "NMOS":
            drop_with_label(page, M_NMOS, inst, W_NMOS, H_NMOS, pin_positions)
        elif inst["type"] == "PMOS":
            drop_with_label(page, M_PMOS, inst, W_PMOS, H_PMOS, pin_positions)
        else:
            drop_with_label(page, M_UNKNOWN, inst, W_UNKNOWN, H_UNKNOWN, pin_positions)

    print("所有器件已放置完成。")

    # 自动连线（仅横/竖线；排除 B 与 VSSA/VSSD）
    draw_net_lines(page, netlist, pin_positions, M_LINE)

    print("\n所有器件与连线已完成。")

if __name__ == "__main__":
    main()

