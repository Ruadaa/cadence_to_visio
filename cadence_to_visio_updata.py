import re
import win32com.client
import math

# === 配置 ===
INPUT_FILE   = r"C:\Users\Administrator\Desktop\cadence_to_visio\inst_info.txt"
NETLIST_FILE = r"C:\Users\Administrator\Desktop\cadence_to_visio\netlist.txt"
STENCIL      = r"C:\Users\Administrator\Desktop\cadence_to_visio\circuit.vss"
SCALE        = 2


EXCLUDED_NETS = {"VDDA", "VSSA", "GNDA"}
EXCLUDED_PINS = {"B"}

# === 统一的器件库 ===
DEVICE_LIBRARY = {
    "NMOS": {
        "inst_prefix": ["NM", "M"],
        "netlist_prefix": ["XNM","XM"],
        "master_name": "NMOS",
        "size": (0.44, 0.59),
        "pins": {
            "D": ( 0.5,  0.5),
            "G": (-0.5, 0.0),
            "S": ( 0.5, -0.5),
            "B": ( 0.0,  0.0),
        }
    },
    "PMOS": {
        "inst_prefix": ["PM"],
        "netlist_prefix": ["XPM"],
        "master_name": "PMOS",
        "size": (0.44, 0.59),
        "pins": {
            "D": ( 0.5, -0.5),
            "G": (-0.5, 0.0),
            "S": ( 0.5,  0.5),
            "B": ( 0.0,  0.0),
        }
    },
    "RES": {
        "inst_prefix": ["R"],
        "netlist_prefix": ["XR"],
        "master_name": "R",
        "size": (0.20, 0.59),
        "pins": {
            "R_up":   (0.0,  0.5),
            "R_down": (0.0, -0.5),
        }
    },
    "Cap": {
        "inst_prefix": ["C"],
        "netlist_prefix": ["CC"],
        "master_name": "C",
        "size": (0.20, 0.59),
        "pins": {
            "C_up":   (0.0,  0.5),
            "C_down": (0.0, -0.5),
        }
    },
    # === 新增 Unknown 器件 ===
    "UNKNOWN": {
        "inst_prefix": [],
        "netlist_prefix": [],
        "master_name": "Unknown",  
        "size": (0.43, 0.43),
        "pins": {
            "P1": (0.0, 0.5),
            "P2": (0.0, -0.5),
            "P3": (0.5, 0),
            "P4": (-0.5, 0),
            "P5": (0.0, 0)
        }
    }
    # 以后你可以自己加新器件
}

def match_device_type(name, from_netlist=False):
    candidates = []
    for dev_type, cfg in DEVICE_LIBRARY.items():
        prefixes = cfg["netlist_prefix"] if from_netlist else cfg["inst_prefix"]
        for p in prefixes:
            candidates.append((len(p), p, dev_type))
    # 按前缀长度从大到小排序
    for _, p, dev_type in sorted(candidates, key=lambda x: -x[0]):
        if name.upper().startswith(p.upper()):
            return dev_type
    return "UNKNOWN"


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

        dev_type = match_device_type(name, from_netlist=False)

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
            raw_name = tokens[0]  # e.g., CC1, CC0

            dev_type = match_device_type(raw_name, from_netlist=True)
            # 期望引脚数：来自 DEVICE_LIBRARY；未知器件则至少 2
            if dev_type in DEVICE_LIBRARY:
                pin_list = list(DEVICE_LIBRARY[dev_type]["pins"].keys())
                pin_count = len(pin_list)
            else:
                pin_count = max(2, len(tokens) - 2)  # 兜底

            if len(tokens) < 1 + pin_count:
                continue  # 行格式不足

            pins = tokens[1:1+pin_count]                # 精确按数量取引脚
            model = tokens[1+pin_count] if len(tokens) > 1+pin_count else ""  # 剩余第一个当模型/值
            # name = raw_name[1:] if raw_name.startswith("X") else raw_name
            name = raw_name[1:]
            # 对未知器件，生成 P1..Pn 引脚名；已知器件用库里的 pin 名
            if dev_type in DEVICE_LIBRARY:
                pin_names = pin_list
            else:
                pin_names = [f"P{i+1}" for i in range(pin_count)]

            pin_map = dict(zip(pin_names, pins))
            devices.append({
                "name": name,
                "type": dev_type,
                "pins": pin_map,
                "model": model
            })
    return devices

# === 放置器件 ===
def drop_with_label(page, master, inst, pin_positions, instances_map):
    dev_type = inst["type"]
    cfg = DEVICE_LIBRARY.get(dev_type, None)
    if not cfg:
        return None
    w, h = cfg["size"]
    cx, cy = inst["xy"]
    name = inst["name"]
    orient = inst["orient"]

    shp = page.Drop(master, cx, cy)
    shp.Text = name
    shp.CellsU("Width").ResultIU  = w
    shp.CellsU("Height").ResultIU = h
    # 文本位置与尺寸
    shp.CellsU("TxtPinX").ResultIU   = shp.CellsU("Width").ResultIU + 0.20
    shp.CellsU("TxtPinY").ResultIU   = shp.CellsU("Height").ResultIU / 2.0
    shp.CellsU("TxtWidth").ResultIU  = 0.6
    shp.CellsU("TxtHeight").ResultIU = 0.2


    apply_orientation(shp, orient)
    instances_map[name] = shp

    # 记录引脚坐标
    for pin, (rx, ry) in cfg["pins"].items():
        pin_positions[f"{name}:{pin}"] = (cx + rx*w, cy + ry*h)

    return shp

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

# === MST 构造 ===
def build_mst(points, candidate_edges=None):
    if candidate_edges is None:
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

# === 绘制连线 ===
def draw_net_lines(page, netlist, pin_positions, instances_map, bboxes):
    net_to_points = {}
    # 收集每个网络的点
    for dev in netlist:
        name = dev["name"]
        dev_type = dev["type"]
        for pin, net in dev["pins"].items():
            if pin.upper() in EXCLUDED_PINS or net.upper() in EXCLUDED_NETS:
                continue
            key = f"{name}:{pin}"
            if key in pin_positions:
                pt = pin_positions[key]
                net_to_points.setdefault(net, []).append((name, dev_type, pin, pt))

    for net, pins in net_to_points.items():
        if len(pins) < 2:
            continue
        coords = [pt for _, _, _, pt in pins]
        edges = build_mst(coords)

        for p1, p2 in edges:
            horiz = abs(p1[1]-p2[1]) < 1e-6
            vert  = abs(p1[0]-p2[0]) < 1e-6
            is_manhattan = horiz or vert

            line = page.Drop(page.Application.ConnectorToolDataObject, 0, 0)
            line.CellsU("BeginX").ResultIU = p1[0]
            line.CellsU("BeginY").ResultIU = p1[1]
            line.CellsU("EndX").ResultIU   = p2[0]
            line.CellsU("EndY").ResultIU   = p2[1]
            if is_manhattan:
                line.CellsU("LineWeight").FormulaU = "1.2 pt"
                # line.CellsU("RouteStyle").FormulaU = "64"  # 正交
            else:
                line.CellsU("LineWeight").FormulaU = "0.6 pt"
                line.CellsU("LinePattern").FormulaU = "2"
                # line.CellsU("RouteStyle").FormulaU = "32"  # 直线

            # === 自动 GlueTo ===
            def find_dev_pin(pt, pins, tol=1e-4):
                tx, ty = pt
                for (dn, dt, pn, (x, y)) in pins:
                    if abs(x - tx) < tol and abs(y - ty) < tol:
                        return dn, dt, pn
                return None, None, None

            dev1, type1, pin1 = find_dev_pin(p1, pins)
            dev2, type2, pin2 = find_dev_pin(p2, pins)

            for dev, dtype, pin, end in [(dev1, type1, pin1, "Begin"), (dev2, type2, pin2, "End")]:
                if dev and dtype in DEVICE_LIBRARY:
                    shape = instances_map[dev]
                    pin_list = list(DEVICE_LIBRARY[dtype]["pins"].keys())
                    if pin in pin_list:
                        idx = pin_list.index(pin) + 1
                        try:
                            conn_x = shape.CellsU(f"Connections.X{idx}")
                            conn_y = shape.CellsU(f"Connections.Y{idx}")
                            line.CellsU(f"{end}X").GlueTo(conn_x)
                            line.CellsU(f"{end}Y").GlueTo(conn_y)
                        except Exception as e:
                            print(f"[Glue] {dev}:{pin} 失败: {e}")

# === 主程序 ===
def main():
    # 启动 Visio
    visio = win32com.client.Dispatch("Visio.Application")
    visio.Visible = True
    doc = visio.Documents.Add("")
    page = visio.ActivePage

    # 打开模具库
    stencil = visio.Documents.OpenEx(STENCIL, 64)
    # 根据 DEVICE_LIBRARY 里的 master_name 建立映射
    masters = {}
    for dev_type, cfg in DEVICE_LIBRARY.items():
        try:
            masters[dev_type] = stencil.Masters(cfg["master_name"])
        except Exception as e:
            print(f"[警告] 模具 {cfg['master_name']} 未找到: {e}")

    # 解析输入文件
    instances = parse_instances(INPUT_FILE)
    netlist   = parse_netlist(NETLIST_FILE)

    pin_positions = {}
    bboxes = {}
    shapes_map = {}

    # 放置器件
    for inst in instances.values():
        dev_type = inst["type"]
        cfg = DEVICE_LIBRARY.get(dev_type, None)
        if not cfg or dev_type not in masters:
            continue
        master = masters[dev_type]
        shp = drop_with_label(page, master, inst, pin_positions, shapes_map)
        if shp:
            w, h = cfg["size"]
            cx, cy = inst["xy"]
            bboxes[inst["name"]] = (cx - w/2, cy - h/2, cx + w/2, cy + h/2)
    print("--------------------------------------------------------------------------")
    print("所有器件已放置完成。")

    print("--------------------------------------------------------------------------")
    print(f"开始自动连线")

    draw_net_lines(page, netlist, pin_positions, shapes_map, bboxes)
    print("--------------------------------------------------------------------------")
    print("所有器件与连线已完成。   粗实线--直连横竖线，细虚线--直角拐角线。   待手动调整......")

    # === 交互式处理虚线 ===
    print("--------------------------------------------------------------------------")
    choice = input("是否要将剩余虚线改为粗实线？(Y/N): ").strip().lower()
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
        print("--------------------------------------------------------------------------")
        print("已将剩余虚线改为实线。")
    else:
        print("--------------------------------------------------------------------------")
        print("保留虚线，不做修改。")


if __name__ == "__main__":
    main()
