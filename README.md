# cadence_to_visio

将 Cadence Virtuoso 原理图导出的器件、网表和走线坐标重建为可编辑的 Microsoft Visio 原理图。

这个仓库现在包含两套互补的使用方式：

- 仓库根目录下的 `cadence_to_visio_v2.py` / `cadence_to_visio_core.py`
  适合你已经手头有 `inst_info`、网表和走线文件时，直接做 Visio 重建。
- `sch2visio_skill/`
  适合配合 `virtuoso-bridge-lite` 做端到端流程：从当前活动的 Virtuoso schematic 远程导出，再自动下载并生成 Visio。

## 主要功能

- 读取 Cadence 导出的器件位置信息、网表和走线坐标
- 在 Visio 中生成可编辑的器件和连线
- 保留尽量接近 Virtuoso 的相对布局与线段形状
- 支持直接使用 `wires.tsv` 与 `netlist.cdl`
- 自动从顶层 `.SUBCKT` 端口识别电源/地网络，例如 `VDD/VSS`、`VCC/GND`、`AVDD/AVSS`
- 对 MOS body wire 做过滤和源极连接恢复，减少错误短支路
- 复用当前 Visio 会话并规避 stencil 文件锁冲突

## 仓库结构

```text
cadence_to_visio_v2.py         直接绘图主入口
cadence_to_visio_core.py       解析、几何计算、网表处理、走线处理核心
circuit.vss                    Visio stencil
export_cdl_current.il          导出当前活动 schematic 的 CDL 网表
export_inst_xy_orient.il       Cadence 器件坐标导出脚本
export_wire_lines_v4.il        Cadence 走线坐标导出脚本
README.md                      中文说明

sch2visio_skill/
  SKILL.md
  scripts/sch2visio.py
  scripts/run_cadence_to_visio.py
  scripts/project/cadence_to_visio_core.py
  scripts/project/cadence_to_visio_v2.py
  references/inputs-and-workflow.md
  assets/visio/circuit.vss
```

## 环境要求

Windows 侧：

- Python 3.10 及以上
- Microsoft Visio
- `pywin32`

可选：

- `openpyxl`

远端 Cadence 侧：

- Cadence Virtuoso
- 可调用的 `sch2visio.il` 或等效导出脚本

推荐配套：

- `virtuoso-bridge-lite`

## 安装

```powershell
pip install pywin32 openpyxl
```

如果你准备使用端到端自动导出流程，推荐把 `virtuoso-bridge-lite` 放在本仓库同级目录，或在运行时通过 `--bridge-root` 显式指定路径。

## 用法一：已有导出文件时，直接生成 Visio

如果你已经从 Virtuoso 手动导出了这些文件：

- `inst_info.txt`
- `netlist.cdl`
- `wires.tsv`

先做 dry-run 验证：

```powershell
python .\sch2visio_skill\scripts\run_cadence_to_visio.py validate `
  --wires C:\path\wires.tsv `
  --netlist C:\path\netlist.cdl `
  --inst-info C:\path\inst_info.txt `
  --cwd C:\path\output
```

再正式绘图：

```powershell
python .\sch2visio_skill\scripts\run_cadence_to_visio.py visio `
  --wires C:\path\wires.tsv `
  --netlist C:\path\netlist.cdl `
  --inst-info C:\path\inst_info.txt `
  --cwd C:\path\output `
  -- --hidden
```

## 用法二：推荐，搭配 virtuoso-bridge-lite 端到端使用

### 1. 准备 bridge

确保 `virtuoso-bridge-lite` 已安装，并且能正常连接到 Virtuoso 所在机器。

### 2. 准备远端导出脚本

在 Cadence 主机上放置一个可被 `load()` 的 `sch2visio.il`，它需要把当前活动 schematic 导出为：

- `netlist.cdl`
- `inst_info.txt`
- `wires.tsv`

如果你只缺少 `netlist.cdl` 导出入口，也可以直接使用仓库根目录新增的 `export_cdl_current.il`。

### 3. 运行自动导出 + 绘图

```powershell
python .\sch2visio_skill\scripts\sch2visio.py `
  --bridge-root ..\virtuoso-bridge-lite `
  --local-dir .\output\demo_visio
```

如果只想检查输入是否正确，不打开 Visio：

```powershell
python .\sch2visio_skill\scripts\sch2visio.py `
  --bridge-root ..\virtuoso-bridge-lite `
  --local-dir .\output\demo_visio `
  --validate-only
```

## 关键参数说明

`sch2visio.py` 常用参数：

- `--bridge-root`
- `--remote-skill`
- `--remote-dir`
- `--local-dir`
- `--output-vsdx`
- `--validate-only`
- `--show-visio`

`run_cadence_to_visio.py` 常用参数：

- `validate`
- `visio`
- `--wires`
- `--netlist`
- `--inst-info`
- `--cwd`

透传给底层绘图脚本的常用附加参数：

- `--skip-nets vdd,vss`
- `--skip-mos-body-nets`
- `--draw-mos-b-wires`
- `--wire-adjust snap-endpoints`
- `--hidden`
- `--preserve-absolute`
- `--flip-y`

## Virtuoso 手动导出参考

导出器件信息：

```lisp
load("/path/to/cadence_to_visio/export_inst_xy_orient.il")
c2vExportInstXYOrient("/path/to/output/inst_info.txt")
```

导出走线：

```lisp
load("/path/to/cadence_to_visio/export_wire_lines_v4.il")
c2vExportWireLinesV4("/path/to/output/wires.tsv")
```

网表建议保存为 `netlist.cdl`。

导出当前活动 schematic 的 CDL：

```lisp
load("/path/to/cadence_to_visio/export_cdl_current.il")
export_cdl_current("/path/to/output/netlist.cdl")
```

## 输入文件说明

- `inst_info.txt`：器件名、坐标、方向、BBox
- `netlist.cdl`：顶层连接关系和器件 pin/net 信息
- `wires.tsv`：精确线段坐标

`wires.tsv` 应至少包含：

- `group_id`
- `seg_id`
- `net`
- `obj_type`
- `layer`
- `purpose`
- `x1`
- `y1`
- `x2`
- `y2`

## 个人信息审查结果

复制进仓库的 `sch2visio_skill` 副本已经处理过这些敏感或环境绑定内容：

- 去掉了本机绝对路径示例，例如 `C:\Users\12398\...`
- 去掉了工作目录示例，例如 `C:\Wangzz\...`
- 去掉了远端个人目录示例，例如 `/home/wzzheng/...`
- 删除了 Visio 运行时锁文件 `assets/visio/~$$circuit.~vss`

保留但不敏感的内容：

- `VB_REMOTE_HOST`
- `VB_REMOTE_USER`

这两个只是环境变量名，不包含你的实际主机名或用户名。

## 注意事项

- 本仓库当前还有其他未提交文件，我这次只会提交 `README.md` 和 `sch2visio_skill/` 相关变更。
- `sch2visio_skill/scripts/run_cadence_to_visio.py` 里仍保留了对 `assets/cadence` 的引用；复制到本仓库后，这部分正好可以直接复用根目录现有的 `.il` 文件。
