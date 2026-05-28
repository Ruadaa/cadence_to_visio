# Cadence to Visio V2.0

将 Cadence/Virtuoso 原理图导出的器件、网表和走线坐标重建到 Microsoft Visio。

目标：尽量保持与 Virtuoso 一样的器件位置和走线位置，生成可编辑的 Visio 原理图，后续在 Visio 中手动微调即可。

![example](example.svg)

## 主要文件

```text
cadence_to_visio_v2.py       主入口，直接运行
cadence_to_visio_core.py     解析、坐标转换和绘图逻辑
device_mapping.toml          器件映射配置（分类规则、stencil master、pin 定义）
device_mapping_editor.py     配置编辑器 GUI
circuit.vss                  Visio stencil
inst_info.txt                器件坐标、方向、BBox
netlist.txt                  CDL 网表
wires.xlsx                   Virtuoso wire 坐标
example.svg                  示例图
old_cadence_to_visio/        旧版本和实验文件归档
```

## 安装

```powershell
pip install pywin32 openpyxl tomlkit
# Python 3.10 需额外安装：pip install tomli
```

需要 Windows + Microsoft Visio。

## 从 Virtuoso 导出

导出器件信息：

```lisp
load("/path/to/cadence_to_visio/export_inst_xy_orient.il")
c2vExportInstXYOrient("/path/to/cadence_to_visio/inst_info.txt")
```

导出走线坐标：

```lisp
load("/path/to/cadence_to_visio/export_wire_lines_v4.il")
c2vExportWireLinesV4("/path/to/cadence_to_visio/wires.tsv")
```

将 `wires.tsv` 用 Excel 另存为 `wires.xlsx`。CDL 网表保存为 `netlist.txt`。

## 运行

准备好 `inst_info.txt`、`netlist.txt`、`wires.xlsx` 后：

```powershell
python .\cadence_to_visio_v2.py
```

只检查输入：

```powershell
python .\cadence_to_visio_v2.py --dry-run
```

默认行为：

- 绘制 node；
- 启用附着；
- 保留 Virtuoso 走线形状；
- 不使用 Visio 自动重路由 connector。

## 常用选项

```powershell
python .\cadence_to_visio_v2.py --no-attach
python .\cadence_to_visio_v2.py --no-draw-nodes
python .\cadence_to_visio_v2.py --draw-mos-b-wires
python .\cadence_to_visio_v2.py --skip-nets vdd,vss
python .\cadence_to_visio_v2.py --wires .\your_wires.xlsx
```

## 器件映射配置

默认内置 NMOS、PMOS、NPN、PNP、R、C、PIN 的分类和映射规则。如果电路中有其他器件（如反相器、传输管、二极管等），可通过 `device_mapping.toml` 自定义映射。

### 配置文件结构

`device_mapping.toml` 包含五个部分：

| 部分 | 说明 |
|------|------|
| `[[classify_rules]]` | 器件分类规则：根据器件名/Cell 名的模式匹配确定器件类型 |
| `[master_candidates]` | Visio stencil master 候选列表：每种器件类型对应的候选 master |
| `[pin_order]` | Pin 顺序：决定网表解析和 Visio connection point 编号 |
| `[pin_hints.XXX]` | Pin 相对位置：器件中心到各 pin 的偏移坐标 |
| `[anchor_type]` | 锚点偏移类型：`left_center`、`bottom_center`、`none` |

分类规则支持通配符：`NM*`（前缀）、`*nmos*`（包含）、`nud18ll_ckt`（精确），逗号分隔多模式。

### 使用配置文件

```powershell
# 默认加载同目录的 device_mapping.toml
python .\cadence_to_visio_v2.py

# 指定配置文件
python .\cadence_to_visio_v2.py --device-mapping .\my_mapping.toml
```

不提供配置文件时使用内置默认值，行为与未配置化前完全一致。

### 配置编辑器 GUI

```powershell
python .\device_mapping_editor.py
```

图形界面支持：
- 分类规则的可视化编辑、排序和匹配测试
- Master 候选列表管理，可从 stencil 扫描可用 master
- Pin 顺序和坐标编辑
- 锚点类型配置
- 配置文件的加载/保存（保留注释格式）

### 示例：添加反相器类型

在 `device_mapping.toml` 中追加：

```toml
[[classify_rules]]
name_pattern = "INV*, I*"
cell_pattern = "inv*, *inverter*"
dev_type = "INV"

[master_candidates]
INV = ["Inverter", "INV", "inv"]

[pin_order]
INV = ["A", "Y"]

[pin_hints.INV]
A = [-0.5, 0.0]
Y = [0.5, 0.0]

[anchor_type]
INV = "none"
```

## 支持器件

支持 NMOS、PMOS、NPN、PNP、R、C、PIN。

NPN/PNP 的 connection points 顺序为 `B, E, C`。MOS 和 BJT 的 Visio anchor 会按方向补偿，使符号位置与 Virtuoso 坐标对齐。
