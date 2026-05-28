"""device_mapping.toml 配置编辑器 GUI。

基于 tkinter + ttk 的可视化编辑器，用于编辑器件映射配置文件。
使用 tomlkit 保留 TOML 文件注释和格式。

用法：
    python device_mapping_editor.py [device_mapping.toml]
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import tomlkit
from tomlkit import table as toml_table, aot as toml_aot

# 尝试导入核心库（用于匹配测试）
try:
    import cadence_to_visio_core as core
except ImportError:
    core = None


# ============================================================
# 数据模型
# ============================================================

@dataclass
class EditorState:
    """编辑器内部状态。"""
    classify_rules: List[Dict[str, str]] = field(default_factory=list)
    master_candidates: Dict[str, List[str]] = field(default_factory=dict)
    pin_order: Dict[str, List[str]] = field(default_factory=dict)
    pin_hints: Dict[str, Dict[str, Tuple[float, float]]] = field(default_factory=dict)
    anchor_type: Dict[str, str] = field(default_factory=dict)
    stencil_masters: List[str] = field(default_factory=list)
    dirty: bool = False
    toml_path: str = ""


# ============================================================
# TOML 读写
# ============================================================

class TomlIO:
    """TOML 文件读写，使用 tomlkit 保留格式和注释。"""

    def __init__(self):
        self.toml_doc: Optional[tomlkit.TOMLDocument] = None

    def load(self, path: str) -> EditorState:
        """从 TOML 文件加载到 EditorState。"""
        with open(path, "r", encoding="utf-8") as f:
            self.toml_doc = tomlkit.load(f)

        state = EditorState()
        state.toml_path = path

        # 分类规则
        state.classify_rules = []
        for rule in self.toml_doc.get("classify_rules", []):
            state.classify_rules.append({
                "name_pattern": rule.get("name_pattern", ""),
                "cell_pattern": rule.get("cell_pattern", ""),
                "dev_type": rule.get("dev_type", ""),
            })

        # master_candidates
        state.master_candidates = {}
        for k, v in self.toml_doc.get("master_candidates", {}).items():
            state.master_candidates[k] = list(v)

        # pin_order
        state.pin_order = {}
        for k, v in self.toml_doc.get("pin_order", {}).items():
            state.pin_order[k] = list(v)

        # pin_hints
        state.pin_hints = {}
        for dev_type, pins in self.toml_doc.get("pin_hints", {}).items():
            state.pin_hints[dev_type] = {}
            for pin, coords in pins.items():
                state.pin_hints[dev_type][pin] = (float(coords[0]), float(coords[1]))

        # anchor_type
        state.anchor_type = {}
        for k, v in self.toml_doc.get("anchor_type", {}).items():
            state.anchor_type[k] = str(v)

        return state

    def save(self, path: str, state: EditorState) -> None:
        """将 EditorState 差异更新到 TOML 文档并保存。"""
        if self.toml_doc is None:
            self.toml_doc = tomlkit.document()

        doc = self.toml_doc

        # classify_rules: 完全重建
        aot = toml_aot()
        for rule in state.classify_rules:
            t = toml_table()
            name_p = rule.get("name_pattern", "")
            cell_p = rule.get("cell_pattern", "")
            if name_p:
                t.add("name_pattern", name_p)
            if cell_p:
                t.add("cell_pattern", cell_p)
            t.add("dev_type", rule.get("dev_type", ""))
            aot.append(t)
        doc["classify_rules"] = aot

        # 其他 section: 原地更新
        self._update_dict_section(doc, "master_candidates", state.master_candidates)
        self._update_dict_section(doc, "pin_order", state.pin_order)
        self._update_anchor_type(doc, state.anchor_type)
        self._update_pin_hints(doc, state.pin_hints)

        with open(path, "w", encoding="utf-8") as f:
            tomlkit.dump(doc, f)

        state.dirty = False
        state.toml_path = path

    def _update_dict_section(self, doc, section_name: str, new_data: dict) -> None:
        """原地更新一个 [section] 的 key-value 映射。"""
        if section_name not in doc:
            doc.add(section_name, toml_table())
        tbl = doc[section_name]

        existing = set(tbl.keys())
        new_keys = set(new_data.keys())

        for key in existing - new_keys:
            del tbl[key]

        for key, value in new_data.items():
            tbl[key] = value

    def _update_anchor_type(self, doc, anchor_type: dict) -> None:
        """更新 anchor_type section。"""
        self._update_dict_section(doc, "anchor_type", anchor_type)

    def _update_pin_hints(self, doc, pin_hints: dict) -> None:
        """原地更新 pin_hints 嵌套 table。"""
        if "pin_hints" not in doc:
            doc.add("pin_hints", toml_table())
        ph = doc["pin_hints"]

        existing = set(ph.keys())
        new_keys = set(pin_hints.keys())

        for key in existing - new_keys:
            del ph[key]

        for dev_type, pins in pin_hints.items():
            if dev_type not in ph:
                ph.add(dev_type, toml_table())
            dt_table = ph[dev_type]

            for old_pin in list(dt_table.keys()):
                del dt_table[old_pin]

            for pin, (dx, dy) in pins.items():
                dt_table.add(pin, [dx, dy])


# ============================================================
# Stencil 扫描
# ============================================================

class StencilScanner:
    """通过 pywin32 COM 扫描 Visio stencil 文件。"""

    @staticmethod
    def is_available() -> bool:
        """检查当前环境是否支持 COM 扫描。"""
        try:
            import win32com.client
            return True
        except ImportError:
            return False

    @staticmethod
    def scan(stencil_path: str) -> List[str]:
        """扫描 .vss 文件，返回所有 master 名称。"""
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError("需要安装 pywin32：pip install pywin32")

        if not os.path.exists(stencil_path):
            raise FileNotFoundError(f"Stencil 文件不存在：{stencil_path}")

        visio = None
        stencil = None
        try:
            visio = win32com.client.Dispatch("Visio.Application")
            visio.Visible = False
            stencil = visio.Documents.OpenEx(os.path.abspath(stencil_path), 64)

            masters = []
            seen = set()
            for idx in range(1, stencil.Masters.Count + 1):
                master = stencil.Masters.Item(idx)
                for attr in ("NameU", "Name"):
                    try:
                        name = str(getattr(master, attr))
                        if name and name.lower() not in seen:
                            seen.add(name.lower())
                            masters.append(name)
                    except Exception:
                        pass

            return sorted(masters)
        finally:
            if stencil:
                try:
                    stencil.Close()
                except Exception:
                    pass
            if visio:
                try:
                    visio.Quit()
                except Exception:
                    pass


# ============================================================
# 主窗口
# ============================================================

class DeviceMappingEditor(tk.Tk):
    """device_mapping.toml 配置编辑器主窗口。"""

    ANCHOR_OPTIONS = ["none", "left_center", "bottom_center"]
    ANCHOR_LABELS = {
        "none": "无偏移",
        "left_center": "左边缘中心（MOS/BJT）",
        "bottom_center": "底部中心（RES/CAP）",
    }

    def __init__(self, toml_path: str = ""):
        super().__init__()
        self.title("Device Mapping 配置编辑器")
        self.geometry("960x680")
        self.minsize(800, 500)

        self.state = EditorState()
        self.toml_io = TomlIO()
        self._current_dev_type: str = ""
        self._suppress_refresh = False

        self._setup_style()
        self._create_toolbar()
        self._create_main_area()
        self._create_status_bar()

        if toml_path:
            self._load_file(toml_path)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ----------------------------------------------------------------
    # 样式
    # ----------------------------------------------------------------

    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Toolbar.TFrame", padding=4)
        style.configure("Status.TLabel", padding=(6, 2))
        style.configure("Treeview", rowheight=24)

    # ----------------------------------------------------------------
    # 工具栏
    # ----------------------------------------------------------------

    def _create_toolbar(self) -> None:
        toolbar = ttk.Frame(self, style="Toolbar.TFrame")
        toolbar.pack(fill="x", side="top")

        ttk.Label(toolbar, text="配置文件:").pack(side="left", padx=(0, 4))
        self._path_var = tk.StringVar()
        self._path_entry = ttk.Entry(toolbar, textvariable=self._path_var, width=50)
        self._path_entry.pack(side="left", fill="x", expand=True, padx=2)

        ttk.Button(toolbar, text="浏览...", command=self._browse_file).pack(side="left", padx=2)
        ttk.Button(toolbar, text="加载", command=self._on_load).pack(side="left", padx=2)
        ttk.Button(toolbar, text="保存", command=self._on_save).pack(side="left", padx=2)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(toolbar, text="新建默认", command=self._new_default).pack(side="left", padx=2)

    # ----------------------------------------------------------------
    # 主区域
    # ----------------------------------------------------------------

    def _create_main_area(self) -> None:
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # 左侧：dev_type 列表
        left_frame = ttk.LabelFrame(paned, text="器件类型 (dev_type)", width=160)
        paned.add(left_frame, weight=0)

        self._dev_type_listbox = tk.Listbox(left_frame, width=16, exportselection=False,
                                            selectmode="browse", activestyle="none")
        self._dev_type_listbox.pack(fill="both", expand=True, padx=4, pady=(4, 2))
        self._dev_type_listbox.bind("<<ListboxSelect>>", self._on_dev_type_select)

        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", padx=4, pady=(2, 4))
        ttk.Button(btn_frame, text="添加类型", command=self._add_dev_type).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="删除类型", command=self._delete_dev_type).pack(side="left", padx=2)

        # 右侧：Notebook
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)

        self._notebook = ttk.Notebook(right_frame)
        self._notebook.pack(fill="both", expand=True)

        # 创建 4 个 tab
        self._rules_frame = ttk.Frame(self._notebook)
        self._master_frame = ttk.Frame(self._notebook)
        self._pin_frame = ttk.Frame(self._notebook)
        self._anchor_frame = ttk.Frame(self._notebook)

        self._notebook.add(self._rules_frame, text="  分类规则  ")
        self._notebook.add(self._master_frame, text="  Master候选  ")
        self._notebook.add(self._pin_frame, text="  Pin配置  ")
        self._notebook.add(self._anchor_frame, text="  锚点类型  ")

        self._build_rules_tab()
        self._build_master_tab()
        self._build_pin_tab()
        self._build_anchor_tab()

    # ----------------------------------------------------------------
    # 状态栏
    # ----------------------------------------------------------------

    def _create_status_bar(self) -> None:
        self._status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self, textvariable=self._status_var, style="Status.TLabel",
                               relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

    def _update_status(self, msg: str = "") -> None:
        if msg:
            self._status_var.set(msg)
        elif self.state.dirty:
            self._status_var.set("● 已修改（未保存）")
        else:
            path = self.state.toml_path or "（未加载文件）"
            self._status_var.set(f"就绪 — {path}")

    # ----------------------------------------------------------------
    # Tab 1: 分类规则
    # ----------------------------------------------------------------

    def _build_rules_tab(self) -> None:
        parent = self._rules_frame

        # 规则列表
        list_frame = ttk.LabelFrame(parent, text="规则列表（按顺序匹配，首条命中生效）")
        list_frame.pack(fill="both", expand=True, padx=6, pady=6)

        cols = ("idx", "dev_type", "name_pattern", "cell_pattern")
        self._rules_tree = ttk.Treeview(list_frame, columns=cols, show="headings",
                                         height=8, selectmode="browse")
        self._rules_tree.heading("idx", text="#")
        self._rules_tree.heading("dev_type", text="dev_type")
        self._rules_tree.heading("name_pattern", text="name_pattern")
        self._rules_tree.heading("cell_pattern", text="cell_pattern")
        self._rules_tree.column("idx", width=36, minwidth=36, anchor="center")
        self._rules_tree.column("dev_type", width=80, minwidth=60)
        self._rules_tree.column("name_pattern", width=200, minwidth=100)
        self._rules_tree.column("cell_pattern", width=280, minwidth=100)
        self._rules_tree.pack(fill="both", expand=True, padx=4, pady=4)
        self._rules_tree.bind("<<TreeviewSelect>>", self._on_rule_select)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self._rules_tree.yview)
        self._rules_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(fill="y", side="right", in_=list_frame)

        # 按钮行
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=6)
        ttk.Button(btn_frame, text="添加", command=self._add_rule).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="删除", command=self._delete_rule).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="上移", command=self._move_rule_up).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="下移", command=self._move_rule_down).pack(side="left", padx=2)

        # 编辑区
        edit_frame = ttk.LabelFrame(parent, text="编辑选中规则")
        edit_frame.pack(fill="x", padx=6, pady=(4, 2))

        row1 = ttk.Frame(edit_frame)
        row1.pack(fill="x", padx=4, pady=2)
        ttk.Label(row1, text="dev_type:").pack(side="left")
        self._rule_dev_type_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self._rule_dev_type_var, width=16).pack(side="left", padx=4)

        row2 = ttk.Frame(edit_frame)
        row2.pack(fill="x", padx=4, pady=2)
        ttk.Label(row2, text="name_pattern:").pack(side="left")
        self._rule_name_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self._rule_name_var, width=40).pack(side="left", padx=4, fill="x", expand=True)

        row3 = ttk.Frame(edit_frame)
        row3.pack(fill="x", padx=4, pady=2)
        ttk.Label(row3, text="cell_pattern:").pack(side="left")
        self._rule_cell_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self._rule_cell_var, width=40).pack(side="left", padx=4, fill="x", expand=True)

        hint = ttk.Label(edit_frame, text="通配符: NM*=前缀, *nmos*=包含, 精确=无星号, 逗号分隔多模式",
                         foreground="gray")
        hint.pack(anchor="w", padx=4, pady=(0, 4))

        ttk.Button(edit_frame, text="应用到选中规则", command=self._apply_rule_edit).pack(padx=4, pady=(0, 4))

        # 匹配测试
        test_frame = ttk.LabelFrame(parent, text="匹配测试")
        test_frame.pack(fill="x", padx=6, pady=(2, 6))

        tr = ttk.Frame(test_frame)
        tr.pack(fill="x", padx=4, pady=4)
        ttk.Label(tr, text="器件名:").pack(side="left")
        self._test_name_var = tk.StringVar()
        ttk.Entry(tr, textvariable=self._test_name_var, width=16).pack(side="left", padx=4)
        ttk.Label(tr, text="Cell名:").pack(side="left")
        self._test_cell_var = tk.StringVar()
        ttk.Entry(tr, textvariable=self._test_cell_var, width=20).pack(side="left", padx=4)
        ttk.Button(tr, text="测试", command=self._test_match).pack(side="left", padx=4)

        self._test_result_var = tk.StringVar()
        ttk.Label(test_frame, textvariable=self._test_result_var, foreground="blue").pack(
            anchor="w", padx=4, pady=(0, 4))

    # ----------------------------------------------------------------
    # Tab 2: Master候选
    # ----------------------------------------------------------------

    def _build_master_tab(self) -> None:
        parent = self._master_frame

        # dev_type 选择
        sel_frame = ttk.Frame(parent)
        sel_frame.pack(fill="x", padx=6, pady=6)
        ttk.Label(sel_frame, text="器件类型:").pack(side="left")
        self._master_dev_type_var = tk.StringVar()
        self._master_dev_type_combo = ttk.Combobox(
            sel_frame, textvariable=self._master_dev_type_var, state="readonly", width=16)
        self._master_dev_type_combo.pack(side="left", padx=4)
        self._master_dev_type_combo.bind("<<ComboboxSelected>>", self._on_master_dev_type_change)

        ttk.Button(sel_frame, text="扫描 stencil...", command=self._scan_stencil).pack(side="right", padx=4)

        # 左右双列表
        lists_frame = ttk.Frame(parent)
        lists_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # 左侧：当前候选
        left = ttk.LabelFrame(lists_frame, text="当前候选（优先级从上到下）")
        left.pack(side="left", fill="both", expand=True, padx=(0, 4))

        self._master_current_tree = ttk.Treeview(left, columns=("master",), show="headings",
                                                   height=10, selectmode="browse")
        self._master_current_tree.heading("master", text="Master 名称")
        self._master_current_tree.pack(fill="both", expand=True, padx=4, pady=4)

        btn_left = ttk.Frame(left)
        btn_left.pack(fill="x", padx=4, pady=(0, 4))
        ttk.Button(btn_left, text="删除", command=self._delete_master).pack(side="left", padx=2)
        ttk.Button(btn_left, text="上移", command=self._move_master_up).pack(side="left", padx=2)
        ttk.Button(btn_left, text="下移", command=self._move_master_down).pack(side="left", padx=2)

        # 右侧：stencil 可用
        right = ttk.LabelFrame(lists_frame, text="Stencil 可用 Master（双击添加）")
        right.pack(side="right", fill="both", expand=True, padx=(4, 0))

        self._master_stencil_tree = ttk.Treeview(right, columns=("master",), show="headings",
                                                   height=10, selectmode="browse")
        self._master_stencil_tree.heading("master", text="Master 名称")
        self._master_stencil_tree.pack(fill="both", expand=True, padx=4, pady=4)
        self._master_stencil_tree.bind("<Double-1>", self._on_stencil_master_dblclick)

        ttk.Button(right, text="刷新列表", command=self._refresh_stencil_list).pack(padx=4, pady=(0, 4))

        # 手动输入
        input_frame = ttk.Frame(parent)
        input_frame.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Label(input_frame, text="手动输入:").pack(side="left")
        self._master_manual_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self._master_manual_var, width=30).pack(side="left", padx=4)
        ttk.Button(input_frame, text="添加", command=self._add_master_manual).pack(side="left", padx=2)

    # ----------------------------------------------------------------
    # Tab 3: Pin配置
    # ----------------------------------------------------------------

    def _build_pin_tab(self) -> None:
        parent = self._pin_frame

        # dev_type 选择
        sel_frame = ttk.Frame(parent)
        sel_frame.pack(fill="x", padx=6, pady=6)
        ttk.Label(sel_frame, text="器件类型:").pack(side="left")
        self._pin_dev_type_var = tk.StringVar()
        self._pin_dev_type_combo = ttk.Combobox(
            sel_frame, textvariable=self._pin_dev_type_var, state="readonly", width=16)
        self._pin_dev_type_combo.pack(side="left", padx=4)
        self._pin_dev_type_combo.bind("<<ComboboxSelected>>", self._on_pin_dev_type_change)

        # Pin 列表
        list_frame = ttk.LabelFrame(parent, text="Pin 列表（顺序决定 connection point 编号）")
        list_frame.pack(fill="both", expand=True, padx=6, pady=(0, 4))

        cols = ("idx", "pin_name", "dx", "dy")
        self._pin_tree = ttk.Treeview(list_frame, columns=cols, show="headings",
                                       height=8, selectmode="browse")
        self._pin_tree.heading("idx", text="#")
        self._pin_tree.heading("pin_name", text="Pin 名称")
        self._pin_tree.heading("dx", text="dx")
        self._pin_tree.heading("dy", text="dy")
        self._pin_tree.column("idx", width=36, minwidth=36, anchor="center")
        self._pin_tree.column("pin_name", width=120, minwidth=80)
        self._pin_tree.column("dx", width=80, minwidth=60)
        self._pin_tree.column("dy", width=80, minwidth=60)
        self._pin_tree.pack(fill="both", expand=True, padx=4, pady=4)
        self._pin_tree.bind("<<TreeviewSelect>>", self._on_pin_select)

        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill="x", padx=4, pady=(0, 4))
        ttk.Button(btn_frame, text="添加", command=self._add_pin).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="删除", command=self._delete_pin).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="上移", command=self._move_pin_up).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="下移", command=self._move_pin_down).pack(side="left", padx=2)

        # 编辑区
        edit_frame = ttk.LabelFrame(parent, text="编辑选中 Pin")
        edit_frame.pack(fill="x", padx=6, pady=(4, 6))

        row = ttk.Frame(edit_frame)
        row.pack(fill="x", padx=4, pady=4)
        ttk.Label(row, text="Pin名:").pack(side="left")
        self._pin_name_var = tk.StringVar()
        ttk.Entry(row, textvariable=self._pin_name_var, width=12).pack(side="left", padx=4)
        ttk.Label(row, text="dx:").pack(side="left")
        self._pin_dx_var = tk.StringVar()
        ttk.Entry(row, textvariable=self._pin_dx_var, width=8).pack(side="left", padx=4)
        ttk.Label(row, text="dy:").pack(side="left")
        self._pin_dy_var = tk.StringVar()
        ttk.Entry(row, textvariable=self._pin_dy_var, width=8).pack(side="left", padx=4)
        ttk.Button(row, text="应用", command=self._apply_pin_edit).pack(side="left", padx=8)

        hint = ttk.Label(edit_frame, text="坐标单位：器件宽高的一半（0.5 = 半宽/半高）",
                         foreground="gray")
        hint.pack(anchor="w", padx=4, pady=(0, 4))

    # ----------------------------------------------------------------
    # Tab 4: 锚点类型
    # ----------------------------------------------------------------

    def _build_anchor_tab(self) -> None:
        parent = self._anchor_frame

        desc = ttk.Label(parent, text=(
            "锚点偏移类型控制 Visio stencil 图形放置时的位置补偿。\n"
            "  left_center — 锚点在左边缘中心（MOS/BJT 等横向器件）\n"
            "  bottom_center — 锚点在底部中心（RES/CAP 等纵向器件）\n"
            "  none — 无偏移"
        ), justify="left", foreground="gray")
        desc.pack(anchor="w", padx=6, pady=6)

        cols = ("dev_type", "anchor_type")
        self._anchor_tree = ttk.Treeview(parent, columns=cols, show="headings",
                                          height=12, selectmode="browse")
        self._anchor_tree.heading("dev_type", text="dev_type")
        self._anchor_tree.heading("anchor_type", text="锚点类型")
        self._anchor_tree.column("dev_type", width=160, minwidth=80)
        self._anchor_tree.column("anchor_type", width=300, minwidth=100)
        self._anchor_tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self._anchor_tree.bind("<<TreeviewSelect>>", self._on_anchor_select)

        edit_frame = ttk.Frame(parent)
        edit_frame.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Label(edit_frame, text="设置选中类型为:").pack(side="left")
        self._anchor_combo_var = tk.StringVar()
        self._anchor_combo = ttk.Combobox(
            edit_frame, textvariable=self._anchor_combo_var,
            values=self.ANCHOR_OPTIONS, state="readonly", width=16)
        self._anchor_combo.pack(side="left", padx=4)
        ttk.Button(edit_frame, text="应用", command=self._apply_anchor).pack(side="left", padx=4)

    # ================================================================
    # 数据加载/保存
    # ================================================================

    def _load_file(self, path: str) -> None:
        """加载 TOML 文件到编辑器。"""
        try:
            self.state = self.toml_io.load(path)
        except Exception as e:
            messagebox.showerror("加载失败", str(e))
            return

        self._path_var.set(path)
        self._suppress_refresh = True
        self._rebuild_dev_types()
        self._refresh_rules_tab()
        self._refresh_master_dev_types()
        self._refresh_anchor_tab()
        self._suppress_refresh = False

        # 选中第一个 dev_type
        if self.state.classify_rules:
            dt = self.state.classify_rules[0].get("dev_type", "")
            if dt:
                self._select_dev_type(dt)

        self._update_status()
        print(f"已加载配置：{len(self.state.classify_rules)} 条规则, "
              f"{len(self.state.master_candidates)} 种器件类型")

    def _on_load(self) -> None:
        path = self._path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先指定有效的配置文件路径")
            return
        self._load_file(path)

    def _on_save(self) -> None:
        path = self._path_var.get().strip()
        if not path:
            path = filedialog.asksaveasfilename(
                defaultextension=".toml",
                filetypes=[("TOML 文件", "*.toml"), ("所有文件", "*.*")],
                initialfile="device_mapping.toml",
            )
            if not path:
                return
            self._path_var.set(path)

        try:
            self._collect_current_edits()
            self.toml_io.save(path, self.state)
            self._update_status()
            messagebox.showinfo("保存成功", f"配置已保存到：{path}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("TOML 文件", "*.toml"), ("所有文件", "*.*")],
            title="选择配置文件",
        )
        if path:
            self._path_var.set(path)

    def _new_default(self) -> None:
        """从核心库的默认配置创建新配置。"""
        if core is None:
            messagebox.showwarning("提示", "无法导入 cadence_to_visio_core，请确保在同一目录运行")
            return

        cfg = core._build_default_device_mapping()
        self.state = EditorState(
            classify_rules=[
                {"name_pattern": " ".join(f"{m}:{v}" for m, v in r.name_patterns),
                 "cell_pattern": " ".join(f"{m}:{v}" for m, v in r.cell_patterns),
                 "dev_type": r.dev_type}
                for r in cfg.classify_rules
            ],
            master_candidates=dict(cfg.master_candidates),
            pin_order=dict(cfg.pin_order),
            pin_hints={k: dict(v) for k, v in cfg.pin_hints.items()},
            anchor_type=dict(cfg.anchor_type),
        )
        # 重建 name_pattern/cell_pattern 为用户可读形式
        for rule_data in cfg.classify_rules:
            entry = {"dev_type": rule_data.dev_type}
            name_parts = []
            for mode, val in rule_data.name_patterns:
                if mode == "startswith":
                    name_parts.append(f"{val}*")
                elif mode == "endswith":
                    name_parts.append(f"*{val}")
                elif mode == "contains":
                    name_parts.append(f"*{val}*")
                else:
                    name_parts.append(val)
            entry["name_pattern"] = ", ".join(name_parts)

            cell_parts = []
            for mode, val in rule_data.cell_patterns:
                if mode == "startswith":
                    cell_parts.append(f"{val}*")
                elif mode == "endswith":
                    cell_parts.append(f"*{val}")
                elif mode == "contains":
                    cell_parts.append(f"*{val}*")
                else:
                    cell_parts.append(val)
            entry["cell_pattern"] = ", ".join(cell_parts)

        # 用更干净的方式重建
        self.state.classify_rules = []
        for rule_data in cfg.classify_rules:
            name_parts = []
            for mode, val in rule_data.name_patterns:
                if mode == "startswith":
                    name_parts.append(f"{val}*")
                elif mode == "endswith":
                    name_parts.append(f"*{val}")
                elif mode == "contains":
                    name_parts.append(f"*{val}*")
                else:
                    name_parts.append(val)
            cell_parts = []
            for mode, val in rule_data.cell_patterns:
                if mode == "startswith":
                    cell_parts.append(f"{val}*")
                elif mode == "endswith":
                    cell_parts.append(f"*{val}")
                elif mode == "contains":
                    cell_parts.append(f"*{val}*")
                else:
                    cell_parts.append(val)
            self.state.classify_rules.append({
                "name_pattern": ", ".join(name_parts),
                "cell_pattern": ", ".join(cell_parts),
                "dev_type": rule_data.dev_type,
            })

        self.state.dirty = True
        self._path_var.set("")

        self._suppress_refresh = True
        self._rebuild_dev_types()
        self._refresh_rules_tab()
        self._refresh_master_dev_types()
        self._refresh_anchor_tab()
        self._suppress_refresh = False

        if self.state.classify_rules:
            dt = self.state.classify_rules[0].get("dev_type", "")
            if dt:
                self._select_dev_type(dt)

        self._update_status("已创建默认配置（未保存）")

    def _collect_current_edits(self) -> None:
        """在保存前，把当前正在编辑的 Entry 内容收集到 state。"""
        # 规则编辑区
        sel = self._rules_tree.selection()
        if sel:
            idx = self._rules_tree.index(sel[0])
            if 0 <= idx < len(self.state.classify_rules):
                dt = self._rule_dev_type_var.get().strip()
                if dt:
                    self.state.classify_rules[idx]["dev_type"] = dt
                self.state.classify_rules[idx]["name_pattern"] = self._rule_name_var.get().strip()
                self.state.classify_rules[idx]["cell_pattern"] = self._rule_cell_var.get().strip()

    # ================================================================
    # dev_type 列表管理
    # ================================================================

    def _all_dev_types(self) -> List[str]:
        """汇总所有出现过的 dev_type。"""
        types = set()
        for rule in self.state.classify_rules:
            if rule.get("dev_type"):
                types.add(rule["dev_type"])
        types.update(self.state.master_candidates.keys())
        types.update(self.state.pin_order.keys())
        types.update(self.state.pin_hints.keys())
        types.update(self.state.anchor_type.keys())
        return sorted(types)

    def _rebuild_dev_types(self) -> None:
        """刷新左侧 dev_type 列表。"""
        current = self._current_dev_type
        self._dev_type_listbox.delete(0, "end")
        for dt in self._all_dev_types():
            self._dev_type_listbox.insert("end", dt)

        # 恢复选中
        if current:
            self._select_dev_type(current)

    def _select_dev_type(self, dev_type: str) -> None:
        """在左侧列表中选中指定 dev_type。"""
        for i in range(self._dev_type_listbox.size()):
            if self._dev_type_listbox.get(i) == dev_type:
                self._dev_type_listbox.selection_clear(0, "end")
                self._dev_type_listbox.selection_set(i)
                self._dev_type_listbox.see(i)
                self._current_dev_type = dev_type
                self._refresh_master_tab()
                self._refresh_pin_tab()
                break

    def _on_dev_type_select(self, event=None) -> None:
        """左侧列表选中事件。"""
        sel = self._dev_type_listbox.curselection()
        if not sel:
            return
        self._current_dev_type = self._dev_type_listbox.get(sel[0])
        if not self._suppress_refresh:
            self._refresh_master_tab()
            self._refresh_pin_tab()

    def _add_dev_type(self) -> None:
        """添加新 dev_type。"""
        dialog = tk.Toplevel(self)
        dialog.title("添加器件类型")
        dialog.geometry("300x100")
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text="dev_type 名称:").pack(padx=10, pady=(10, 4))
        var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=var, width=20)
        entry.pack(padx=10)
        entry.focus_set()

        def confirm():
            name = var.get().strip().upper()
            if not name:
                return
            existing = self._all_dev_types()
            if name in existing:
                messagebox.showwarning("提示", f"器件类型 '{name}' 已存在", parent=dialog)
                return

            # 在各 section 中创建空条目
            self.state.master_candidates.setdefault(name, [])
            self.state.pin_order.setdefault(name, [])
            self.state.pin_hints.setdefault(name, {})
            self.state.anchor_type.setdefault(name, "none")
            self.state.dirty = True

            self._rebuild_dev_types()
            self._select_dev_type(name)
            self._refresh_master_dev_types()
            self._refresh_anchor_tab()
            dialog.destroy()

        ttk.Button(dialog, text="确定", command=confirm).pack(pady=10)
        entry.bind("<Return>", lambda e: confirm())

    def _delete_dev_type(self) -> None:
        """删除选中的 dev_type 及其所有配置。"""
        if not self._current_dev_type:
            return
        dt = self._current_dev_type
        if not messagebox.askyesno("确认删除", f"删除器件类型 '{dt}' 的所有配置？"):
            return

        self.state.classify_rules = [r for r in self.state.classify_rules if r.get("dev_type") != dt]
        self.state.master_candidates.pop(dt, None)
        self.state.pin_order.pop(dt, None)
        self.state.pin_hints.pop(dt, None)
        self.state.anchor_type.pop(dt, None)
        self.state.dirty = True

        self._current_dev_type = ""
        self._rebuild_dev_types()
        self._refresh_rules_tab()
        self._refresh_master_dev_types()
        self._refresh_anchor_tab()
        self._update_status()

    # ================================================================
    # 分类规则交互
    # ================================================================

    def _refresh_rules_tab(self) -> None:
        """刷新分类规则 Treeview。"""
        self._rules_tree.delete(*self._rules_tree.get_children())
        for i, rule in enumerate(self.state.classify_rules):
            self._rules_tree.insert("", "end", iid=str(i), values=(
                i + 1,
                rule.get("dev_type", ""),
                rule.get("name_pattern", ""),
                rule.get("cell_pattern", ""),
            ))

    def _on_rule_select(self, event=None) -> None:
        sel = self._rules_tree.selection()
        if not sel:
            return
        idx = self._rules_tree.index(sel[0])
        rule = self.state.classify_rules[idx]
        self._rule_dev_type_var.set(rule.get("dev_type", ""))
        self._rule_name_var.set(rule.get("name_pattern", ""))
        self._rule_cell_var.set(rule.get("cell_pattern", ""))

    def _add_rule(self) -> None:
        self.state.classify_rules.append({
            "name_pattern": "",
            "cell_pattern": "",
            "dev_type": "NEW_TYPE",
        })
        self.state.dirty = True
        self._refresh_rules_tab()
        self._rebuild_dev_types()
        self._refresh_master_dev_types()
        self._refresh_anchor_tab()
        # 选中新添加的规则
        last = str(len(self.state.classify_rules) - 1)
        self._rules_tree.selection_set(last)
        self._rules_tree.see(last)
        self._update_status()

    def _delete_rule(self) -> None:
        sel = self._rules_tree.selection()
        if not sel:
            return
        idx = self._rules_tree.index(sel[0])
        del self.state.classify_rules[idx]
        self.state.dirty = True
        self._refresh_rules_tab()
        self._rebuild_dev_types()
        self._update_status()

    def _move_rule_up(self) -> None:
        sel = self._rules_tree.selection()
        if not sel:
            return
        idx = self._rules_tree.index(sel[0])
        if idx == 0:
            return
        rules = self.state.classify_rules
        rules[idx], rules[idx - 1] = rules[idx - 1], rules[idx]
        self.state.dirty = True
        self._refresh_rules_tab()
        self._rules_tree.selection_set(str(idx - 1))

    def _move_rule_down(self) -> None:
        sel = self._rules_tree.selection()
        if not sel:
            return
        idx = self._rules_tree.index(sel[0])
        rules = self.state.classify_rules
        if idx >= len(rules) - 1:
            return
        rules[idx], rules[idx + 1] = rules[idx + 1], rules[idx]
        self.state.dirty = True
        self._refresh_rules_tab()
        self._rules_tree.selection_set(str(idx + 1))

    def _apply_rule_edit(self) -> None:
        sel = self._rules_tree.selection()
        if not sel:
            return
        idx = self._rules_tree.index(sel[0])
        dt = self._rule_dev_type_var.get().strip()
        if not dt:
            messagebox.showwarning("提示", "dev_type 不能为空")
            return
        self.state.classify_rules[idx]["dev_type"] = dt
        self.state.classify_rules[idx]["name_pattern"] = self._rule_name_var.get().strip()
        self.state.classify_rules[idx]["cell_pattern"] = self._rule_cell_var.get().strip()
        self.state.dirty = True
        self._refresh_rules_tab()
        self._rules_tree.selection_set(str(idx))
        self._rebuild_dev_types()
        self._refresh_master_dev_types()
        self._refresh_anchor_tab()
        self._update_status()

    def _test_match(self) -> None:
        """测试器件名/Cell 名匹配当前规则。"""
        name = self._test_name_var.get().strip()
        cell = self._test_cell_var.get().strip()
        if not name and not cell:
            self._test_result_var.set("请输入器件名或 Cell 名")
            return

        if core is not None:
            # 使用核心库的匹配函数
            for i, rule_data in enumerate(self.state.classify_rules):
                name_patterns = core._compile_pattern(rule_data.get("name_pattern", ""))
                cell_patterns = core._compile_pattern(rule_data.get("cell_pattern", ""))

                name_hit = bool(name_patterns) and core._match_pattern(name, name_patterns)
                cell_hit = bool(cell_patterns) and core._match_pattern(cell, cell_patterns)

                if name_patterns and cell_patterns:
                    if name_hit or cell_hit:
                        self._test_result_var.set(
                            f"→ {rule_data['dev_type']}（命中规则 #{i + 1}）")
                        return
                elif name_patterns and name_hit:
                    self._test_result_var.set(
                        f"→ {rule_data['dev_type']}（命中规则 #{i + 1}: name_pattern）")
                    return
                elif cell_patterns and cell_hit:
                    self._test_result_var.set(
                        f"→ {rule_data['dev_type']}（命中规则 #{i + 1}: cell_pattern）")
                    return

            self._test_result_var.set("→ UNKNOWN（无匹配规则）")
        else:
            self._test_result_var.set("（需要 cadence_to_visio_core 支持匹配测试）")

    # ================================================================
    # Master候选交互
    # ================================================================

    def _refresh_master_dev_types(self) -> None:
        """刷新 Master 候选和 Pin 配置的 dev_type 下拉框。"""
        types = self._all_dev_types()
        self._master_dev_type_combo["values"] = types
        self._pin_dev_type_combo["values"] = types

    def _on_master_dev_type_change(self, event=None) -> None:
        dt = self._master_dev_type_var.get()
        if dt:
            self._select_dev_type(dt)

    def _refresh_master_tab(self) -> None:
        """刷新当前 dev_type 的 master 候选列表。"""
        dt = self._current_dev_type
        self._master_dev_type_var.set(dt)

        self._master_current_tree.delete(*self._master_current_tree.get_children())
        candidates = self.state.master_candidates.get(dt, [])
        for i, master_name in enumerate(candidates):
            self._master_current_tree.insert("", "end", iid=str(i), values=(master_name,))

    def _delete_master(self) -> None:
        sel = self._master_current_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._master_current_tree.index(sel[0])
        dt = self._current_dev_type
        if dt in self.state.master_candidates and idx < len(self.state.master_candidates[dt]):
            del self.state.master_candidates[dt][idx]
            self.state.dirty = True
            self._refresh_master_tab()
            self._update_status()

    def _move_master_up(self) -> None:
        sel = self._master_current_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._master_current_tree.index(sel[0])
        dt = self._current_dev_type
        mc = self.state.master_candidates.get(dt, [])
        if idx == 0 or idx >= len(mc):
            return
        mc[idx], mc[idx - 1] = mc[idx - 1], mc[idx]
        self.state.dirty = True
        self._refresh_master_tab()
        self._master_current_tree.selection_set(str(idx - 1))

    def _move_master_down(self) -> None:
        sel = self._master_current_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._master_current_tree.index(sel[0])
        dt = self._current_dev_type
        mc = self.state.master_candidates.get(dt, [])
        if idx >= len(mc) - 1:
            return
        mc[idx], mc[idx + 1] = mc[idx + 1], mc[idx]
        self.state.dirty = True
        self._refresh_master_tab()
        self._master_current_tree.selection_set(str(idx + 1))

    def _add_master_manual(self) -> None:
        name = self._master_manual_var.get().strip()
        if not name or not self._current_dev_type:
            return
        dt = self._current_dev_type
        self.state.master_candidates.setdefault(dt, [])
        if name not in self.state.master_candidates[dt]:
            self.state.master_candidates[dt].append(name)
            self.state.dirty = True
            self._refresh_master_tab()
            self._update_status()
        self._master_manual_var.set("")

    def _on_stencil_master_dblclick(self, event=None) -> None:
        """双击 stencil master 添加到当前候选。"""
        sel = self._master_stencil_tree.selection()
        if not sel or not self._current_dev_type:
            return
        item = self._master_stencil_tree.item(sel[0])
        master_name = item["values"][0]
        dt = self._current_dev_type
        self.state.master_candidates.setdefault(dt, [])
        if master_name not in self.state.master_candidates[dt]:
            self.state.master_candidates[dt].append(master_name)
            self.state.dirty = True
            self._refresh_master_tab()
            self._update_status()

    def _scan_stencil(self) -> None:
        """扫描 stencil 文件获取 master 列表。"""
        if not StencilScanner.is_available():
            messagebox.showinfo("提示",
                                "Stencil 扫描需要 Windows + Visio + pywin32。\n"
                                "当前环境不支持，请手动输入 Master 名称。")
            return

        stencil_path = filedialog.askopenfilename(
            filetypes=[("Visio Stencil", "*.vss *.vssx"), ("所有文件", "*.*")],
            title="选择 Stencil 文件",
            initialdir=os.path.dirname(self.state.toml_path) or ".",
        )
        if not stencil_path:
            return

        self._update_status("正在扫描 stencil...")
        self.update_idletasks()

        try:
            masters = StencilScanner.scan(stencil_path)
            self.state.stencil_masters = masters
            self._refresh_stencil_list()
            self._update_status(f"扫描到 {len(masters)} 个 Master")
        except Exception as e:
            messagebox.showerror("扫描失败", str(e))
            self._update_status()

    def _refresh_stencil_list(self) -> None:
        """刷新 stencil 可用 master 列表。"""
        self._master_stencil_tree.delete(*self._master_stencil_tree.get_children())
        for name in self.state.stencil_masters:
            self._master_stencil_tree.insert("", "end", values=(name,))

    # ================================================================
    # Pin配置交互
    # ================================================================

    def _on_pin_dev_type_change(self, event=None) -> None:
        dt = self._pin_dev_type_var.get()
        if dt:
            self._select_dev_type(dt)

    def _refresh_pin_tab(self) -> None:
        """刷新当前 dev_type 的 pin 列表。"""
        dt = self._current_dev_type
        self._pin_dev_type_var.set(dt)

        self._pin_tree.delete(*self._pin_tree.get_children())
        pin_order = self.state.pin_order.get(dt, [])
        pin_hints = self.state.pin_hints.get(dt, {})

        for i, pin_name in enumerate(pin_order):
            dx, dy = pin_hints.get(pin_name, (0.0, 0.0))
            self._pin_tree.insert("", "end", iid=str(i), values=(i + 1, pin_name, dx, dy))

    def _on_pin_select(self, event=None) -> None:
        sel = self._pin_tree.selection()
        if not sel:
            return
        idx = self._pin_tree.index(sel[0])
        dt = self._current_dev_type
        pin_order = self.state.pin_order.get(dt, [])
        if idx >= len(pin_order):
            return
        pin_name = pin_order[idx]
        hints = self.state.pin_hints.get(dt, {})
        dx, dy = hints.get(pin_name, (0.0, 0.0))
        self._pin_name_var.set(pin_name)
        self._pin_dx_var.set(str(dx))
        self._pin_dy_var.set(str(dy))

    def _add_pin(self) -> None:
        dt = self._current_dev_type
        if not dt:
            return
        self.state.pin_order.setdefault(dt, [])
        self.state.pin_hints.setdefault(dt, {})
        pin_idx = len(self.state.pin_order[dt]) + 1
        new_pin = f"P{pin_idx}"
        self.state.pin_order[dt].append(new_pin)
        self.state.pin_hints[dt][new_pin] = (0.0, 0.0)
        self.state.dirty = True
        self._refresh_pin_tab()
        self._update_status()

    def _delete_pin(self) -> None:
        sel = self._pin_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._pin_tree.index(sel[0])
        dt = self._current_dev_type
        pin_order = self.state.pin_order.get(dt, [])
        if idx >= len(pin_order):
            return
        pin_name = pin_order[idx]
        del pin_order[idx]
        self.state.pin_hints.get(dt, {}).pop(pin_name, None)
        self.state.dirty = True
        self._refresh_pin_tab()
        self._update_status()

    def _move_pin_up(self) -> None:
        sel = self._pin_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._pin_tree.index(sel[0])
        dt = self._current_dev_type
        po = self.state.pin_order.get(dt, [])
        if idx == 0 or idx >= len(po):
            return
        po[idx], po[idx - 1] = po[idx - 1], po[idx]
        self.state.dirty = True
        self._refresh_pin_tab()
        self._pin_tree.selection_set(str(idx - 1))

    def _move_pin_down(self) -> None:
        sel = self._pin_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._pin_tree.index(sel[0])
        dt = self._current_dev_type
        po = self.state.pin_order.get(dt, [])
        if idx >= len(po) - 1:
            return
        po[idx], po[idx + 1] = po[idx + 1], po[idx]
        self.state.dirty = True
        self._refresh_pin_tab()
        self._pin_tree.selection_set(str(idx + 1))

    def _apply_pin_edit(self) -> None:
        sel = self._pin_tree.selection()
        if not sel or not self._current_dev_type:
            return
        idx = self._pin_tree.index(sel[0])
        dt = self._current_dev_type
        po = self.state.pin_order.get(dt, [])
        if idx >= len(po):
            return

        old_pin = po[idx]
        new_pin = self._pin_name_var.get().strip()
        if not new_pin:
            messagebox.showwarning("提示", "Pin 名称不能为空")
            return

        try:
            dx = float(self._pin_dx_var.get())
            dy = float(self._pin_dy_var.get())
        except ValueError:
            messagebox.showwarning("提示", "dx/dy 必须为数字")
            return

        # 如果 pin 名称变了，更新 hints 中的 key
        ph = self.state.pin_hints.get(dt, {})
        if new_pin != old_pin:
            if old_pin in ph:
                coords = ph.pop(old_pin)
                ph[new_pin] = coords
            else:
                ph[new_pin] = (dx, dy)
        else:
            ph[new_pin] = (dx, dy)

        po[idx] = new_pin
        self.state.dirty = True
        self._refresh_pin_tab()
        self._pin_tree.selection_set(str(idx))
        self._update_status()

    # ================================================================
    # 锚点类型交互
    # ================================================================

    def _refresh_anchor_tab(self) -> None:
        """刷新锚点类型列表。"""
        self._anchor_tree.delete(*self._anchor_tree.get_children())
        all_types = self._all_dev_types()
        for dt in all_types:
            at = self.state.anchor_type.get(dt, "none")
            label = self.ANCHOR_LABELS.get(at, at)
            self._anchor_tree.insert("", "end", iid=dt, values=(dt, label))

    def _on_anchor_select(self, event=None) -> None:
        sel = self._anchor_tree.selection()
        if not sel:
            return
        dt = sel[0]
        at = self.state.anchor_type.get(dt, "none")
        self._anchor_combo_var.set(at)

    def _apply_anchor(self) -> None:
        sel = self._anchor_tree.selection()
        if not sel:
            return
        dt = sel[0]
        value = self._anchor_combo_var.get()
        if value:
            self.state.anchor_type[dt] = value
            self.state.dirty = True
            self._refresh_anchor_tab()
            self._anchor_tree.selection_set(dt)
            self._update_status()

    # ================================================================
    # 窗口关闭
    # ================================================================

    def _on_close(self) -> None:
        if self.state.dirty:
            result = messagebox.askyesnocancel(
                "确认关闭",
                "配置已修改但未保存。\n是否保存后再关闭？"
            )
            if result is None:  # 取消
                return
            if result:  # 是：保存
                self._on_save()
        self.destroy()


# ============================================================
# 入口
# ============================================================

def main():
    toml_path = ""
    if len(sys.argv) > 1:
        toml_path = sys.argv[1]

    if not toml_path:
        default = os.path.join(os.path.dirname(__file__) or ".", "device_mapping.toml")
        if os.path.exists(default):
            toml_path = default

    app = DeviceMappingEditor(toml_path)
    app.mainloop()


if __name__ == "__main__":
    main()
