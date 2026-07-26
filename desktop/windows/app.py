import contextlib
import datetime as dt
import io
import re
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


ROOT = Path(__file__).resolve().parents[2]
SHARED_ROOT = ROOT / "shared" / "fulfillment"
DEFAULT_OUTPUT_BASE = Path.home() / "Desktop" / "Codex Folder" / "DropShipZone" / "Codex发货记录"

DEFAULT_DOWNLOADS = Path.home() / "Downloads"
DEFAULT_TEMPLATE_DIR = Path.home() / "Desktop" / "Dropshipzone" / "各种模板"
DEFAULT_SKU_DIR = Path.home() / "Desktop" / "Dropshipzone"

if str(SHARED_ROOT) not in sys.path:
    sys.path.insert(0, str(SHARED_ROOT))

import generate_dp_shipment_upload
import generate_sf_international_template
import generate_yunexpress_template
import sort_yunexpress_labels_by_sku
import split_sf_labels_by_sender


SCRIPT_MODULES = {
    "generate_yunexpress_template.py": generate_yunexpress_template,
    "generate_sf_international_template.py": generate_sf_international_template,
    "generate_dp_shipment_upload.py": generate_dp_shipment_upload,
    "sort_yunexpress_labels_by_sku.py": sort_yunexpress_labels_by_sku,
    "split_sf_labels_by_sender.py": split_sf_labels_by_sender,
}


def latest_file(folder: Path, pattern: str) -> str:
    files = [p for p in folder.glob(pattern) if p.is_file() and not p.name.startswith("~$")]
    if not files:
        return ""
    return str(max(files, key=lambda p: p.stat().st_mtime))


def latest_file_across(candidates: list[tuple[Path, str]]) -> str:
    files = [
        path
        for folder, pattern in candidates
        for path in folder.glob(pattern)
        if path.is_file() and not path.name.startswith("~$")
    ]
    return str(max(files, key=lambda path: path.stat().st_mtime)) if files else ""


def date_folder(date_text: str) -> Path:
    year, month, day = [int(part) for part in date_text.split("-")]
    return DEFAULT_OUTPUT_BASE / f"{year}年{month}月{day}日"


class Tooltip:
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tip: tk.Toplevel | None = None
        self.widget.bind("<Enter>", self.show, add="+")
        self.widget.bind("<Leave>", self.hide, add="+")
        self.widget.bind("<ButtonPress>", self.hide, add="+")

    def show(self, _event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip,
            text=self.text,
            bg="#111827",
            fg="#f8fafc",
            padx=11,
            pady=8,
            justify="left",
            wraplength=360,
            relief="flat",
            font=("Microsoft YaHei UI", 9),
        ).pack()

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


class App:
    COLORS = {
        "bg": "#eaf1f8",
        "ink": "#162033",
        "muted": "#68758a",
        "soft": "#f4f8fd",
        "surface": "#fbfdff",
        "line": "#d9e3ef",
        "charcoal": "#fbfdff",
        "charcoal_2": "#e8f2ff",
        "accent": "#0a84ff",
        "accent_dark": "#0064d8",
        "accent_soft": "#dcecff",
        "button_soft": "#edf3fa",
        "log_bg": "#f4f8fd",
        "log_text": "#344258",
        "success": "#248a5b",
    }

    FILE_LABELS = {
        "dp_orders": "DP订单CSV",
        "sku": "SKU商品库",
        "yun_template": "云途标准模板",
        "sf_template": "顺丰国际标准模板",
        "yun_orders": "云途订单信息",
        "sf_orders": "顺丰国际订单数据",
        "dp_template": "DP发货模板",
        "label_zip": "云途面单ZIP",
        "sf_label_pdf": "顺丰国际面单PDF",
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DP International Fulfillment To Au")
        self.root.minsize(980, 570)
        self.root.configure(bg=self.COLORS["bg"])

        today = dt.date.today().isoformat()
        self.vars = {
            "date": tk.StringVar(value=today),
            "dp_orders": tk.StringVar(value=latest_file(DEFAULT_DOWNLOADS, "Exported_Orders_*.csv")),
            "sku": tk.StringVar(value=latest_file(DEFAULT_SKU_DIR, "*SKU*.xlsx")),
            "yun_template": tk.StringVar(value=latest_file(DEFAULT_TEMPLATE_DIR, "*CN0C021578*.xlsx")),
            "sf_template": tk.StringVar(
                value=latest_file_across(
                    [
                        (DEFAULT_DOWNLOADS, "BatchImportOrders_*.xlsm"),
                        (DEFAULT_TEMPLATE_DIR, "BatchImportOrders_*.xlsm"),
                    ]
                )
            ),
            "yun_orders": tk.StringVar(value=latest_file(DEFAULT_DOWNLOADS, "订单信息_*.xlsx")),
            "sf_orders": tk.StringVar(value=latest_file(DEFAULT_DOWNLOADS, "顺丰订单数据*.xls*")),
            "dp_template": tk.StringVar(value=latest_file(DEFAULT_TEMPLATE_DIR, "*Shipment_sample.xlsx")),
            "label_zip": tk.StringVar(value=latest_file(DEFAULT_DOWNLOADS, "运单标签*.zip")),
            "sf_label_pdf": tk.StringVar(value=latest_file(DEFAULT_DOWNLOADS, "labelSF*.pdf")),
            "shipping_platform": tk.StringVar(value="yunexpress"),
            "sf_business_type": tk.StringVar(value="请选择业务类型"),
            "sf_battery": tk.StringVar(value="否"),
        }
        self.output_dir = tk.StringVar(value=str(date_folder(today)))
        self.status_text = tk.StringVar(value="就绪")
        self.current_step = 0
        self.running = False
        self.last_result_dir: Path | None = None
        self.nav_buttons: list[tk.Button] = []
        self.action_buttons: list[tk.Button] = []
        self.result_buttons: list[tk.Button] = []
        self.platform_buttons: dict[str, tk.Button] = {}

        self.config_files = [
            ("SKU商品库", "sku", [("Excel", "*.xlsx")]),
            ("云途标准模板", "yun_template", [("Excel", "*.xlsx")]),
            ("顺丰国际标准模板", "sf_template", [("Excel 宏工作簿", "*.xlsm")]),
            ("DP发货模板", "dp_template", [("Excel", "*.xlsx")]),
        ]

        self.steps = [
            {
                "index": "01",
                "short": "云途上传",
                "nav": "DP订单 -> 云途模板",
                "title": "生成云途上传模板",
                "summary": "从 DP 订单 CSV 生成云途批量寄件 Excel。",
                "when": "刚从 DP 后台下载订单 CSV 后使用。",
                "files": [("本次 DP 订单 CSV", "dp_orders", [("CSV", "*.csv")])],
                "config_note": "使用固定配置：SKU商品库、云途标准模板。",
                "button": "开始生成",
                "after": "生成后到云途后台上传 Excel。创建寄件订单成功后，下载“订单信息”，再做第2步。",
                "command": self.run_yun,
            },
            {
                "index": "02",
                "short": "DP回填",
                "nav": "云途订单 -> DP上传",
                "title": "生成 DP 发货回填模板",
                "summary": "把云途运单号整理成 DP 后台可上传的发货文件。",
                "when": "云途已生成寄件订单，并已下载“订单信息”Excel 后使用。",
                "files": [
                    ("本次云途订单信息 Excel", "yun_orders", [("Excel", "*.xlsx")]),
                ],
                "config_note": "使用固定配置：DP发货模板。",
                "button": "开始生成",
                "after": "生成后上传到 DropShipZone 后台。成功后下载云途面单 ZIP，再做第3步。",
                "command": self.run_dp,
            },
            {
                "index": "03",
                "short": "面单分拣",
                "nav": "面单ZIP -> 供应商包",
                "title": "按 SKU 分拣面单",
                "summary": "把云途面单 ZIP 拆成供应商可直接接收的 SKU 文件夹和压缩包。",
                "when": "云途面单 ZIP 已下载后使用。",
                "files": [
                    ("本次云途面单 ZIP", "label_zip", [("ZIP", "*.zip")]),
                ],
                "config_note": "自动沿用第1步 DP 订单 CSV 和第2步云途订单信息；使用固定配置：SKU商品库。",
                "button": "开始分拣",
                "after": "每个 SKU 文件夹内会有 PDF 和供应商 ZIP；总目录保留汇总表和校验报告。",
                "command": self.run_labels,
            },
        ]

        self.build_layout()
        self.select_step(0)
        self.fit_initial_window()
        self.write("工具已准备好。第一次处理订单时，请从第1步开始。")

    def fit_initial_window(self):
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        target_w = min(1120, max(980, screen_w - 90))
        target_h = min(620, max(570, screen_h - 130))
        req_w = min(max(self.root.winfo_reqwidth() + 24, target_w), screen_w - 40)
        req_h = min(max(self.root.winfo_reqheight() + 24, target_h), screen_h - 70)
        x = max(0, (screen_w - req_w) // 2)
        y = max(0, (screen_h - req_h) // 2)
        self.root.geometry(f"{req_w}x{req_h}+{x}+{y}")

    def build_layout(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_rowconfigure(3, weight=1)

        self.build_topbar()
        self.build_step_nav()
        self.build_main_area()
        self.build_log_area()

    def build_topbar(self):
        top = tk.Frame(
            self.root,
            bg=self.COLORS["surface"],
            padx=18,
            pady=11,
            highlightthickness=1,
            highlightbackground=self.COLORS["line"],
        )
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        brand = tk.Frame(top, bg=self.COLORS["surface"])
        brand.grid(row=0, column=0, sticky="w")
        tk.Label(
            brand,
            text="DP 国际发货工具",
            bg=self.COLORS["surface"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            brand,
            text="DropShipZone  ·  YunExpress  ·  SF International",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
            font=("Segoe UI", 8),
        ).grid(row=1, column=0, sticky="w", pady=(1, 0))

        controls = tk.Frame(top, bg=self.COLORS["surface"])
        controls.grid(row=0, column=1, sticky="e")

        date_label = tk.Frame(controls, bg=self.COLORS["surface"])
        date_label.grid(row=0, column=0, sticky="e", padx=(0, 6))
        tk.Label(
            date_label,
            text="发货日期",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
            font=("Microsoft YaHei UI", 9, "bold"),
        ).pack(side="left")
        self.info_dot(
            date_label,
            "默认使用电脑当前日期。需要按其他发货日期归档时，可以手动改成 YYYY-MM-DD。",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
        ).pack(side="left", padx=(4, 0))
        date_entry = self.entry(controls, self.vars["date"], width=13)
        date_entry.grid(row=0, column=1, sticky="e", padx=(0, 10))
        date_entry.bind("<FocusOut>", lambda _event: self.refresh_output_dir())

        self.secondary_button(controls, "自动识别", self.autofill_latest, compact=True).grid(row=0, column=2, padx=(0, 6))
        self.secondary_button(controls, "固定配置", self.open_config_dialog, compact=True).grid(row=0, column=3, padx=(0, 6))
        self.secondary_button(controls, "输出目录", self.open_output_dir, compact=True).grid(row=0, column=4)

        status = tk.Frame(controls, bg=self.COLORS["charcoal_2"], padx=10, pady=5)
        status.grid(row=0, column=5, sticky="e", padx=(8, 0))
        status_dot = tk.Canvas(
            status,
            width=9,
            height=9,
            bg=self.COLORS["charcoal_2"],
            highlightthickness=0,
            bd=0,
        )
        status_dot.create_oval(2, 2, 7, 7, fill=self.COLORS["accent"], outline="")
        status_dot.pack(side="left", padx=(0, 6))
        tk.Label(
            status,
            textvariable=self.status_text,
            bg=self.COLORS["charcoal_2"],
            fg=self.COLORS["accent_dark"],
            font=("Microsoft YaHei UI", 9, "bold"),
        ).pack(side="left")

    def build_step_nav(self):
        nav_wrap = tk.Frame(self.root, bg=self.COLORS["bg"], padx=18, pady=10)
        nav_wrap.grid(row=1, column=0, sticky="ew")
        nav = tk.Frame(
            nav_wrap,
            bg="#dfe8f3",
            padx=4,
            pady=4,
            highlightthickness=1,
            highlightbackground="#d4dfec",
        )
        nav.pack(fill="x")
        for i in range(3):
            nav.grid_columnconfigure(i, weight=1, uniform="step")

        for index, step in enumerate(self.steps):
            button = tk.Button(
                nav,
                text=f"{step['index']}  {step['short']}  {step['nav']}",
                command=lambda idx=index: self.select_step(idx),
                anchor="w",
                justify="left",
                relief="flat",
                bd=0,
                padx=15,
                pady=8,
                cursor="hand2",
                font=("Microsoft YaHei UI", 9, "bold"),
            )
            button.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 6, 0 if index == 2 else 6))
            self.nav_buttons.append(button)

    def build_main_area(self):
        self.main = tk.Frame(self.root, bg=self.COLORS["bg"], padx=18)
        self.main.grid(row=2, column=0, sticky="ew")
        self.main.grid_columnconfigure(0, weight=1)

        self.step_panel = self.card(self.main, pad=16)
        self.step_panel.grid(row=0, column=0, sticky="ew")
        self.step_panel.grid_columnconfigure(0, weight=1)

    def build_log_area(self):
        log_frame = tk.Frame(self.root, bg=self.COLORS["bg"], padx=18, pady=0)
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(8, 12))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        bar = tk.Frame(log_frame, bg=self.COLORS["bg"])
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        bar.grid_columnconfigure(0, weight=1)
        tk.Label(
            bar,
            text="运行日志",
            bg=self.COLORS["bg"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            bar,
            textvariable=self.output_dir,
            bg=self.COLORS["bg"],
            fg=self.COLORS["muted"],
            font=("Microsoft YaHei UI", 8),
        ).grid(row=0, column=1, sticky="e")

        self.log = tk.Text(
            log_frame,
            height=3,
            wrap="word",
            relief="flat",
            bd=0,
            bg=self.COLORS["log_bg"],
            fg=self.COLORS["log_text"],
            insertbackground=self.COLORS["log_text"],
            font=("Cascadia Mono", 9),
            padx=12,
            pady=7,
        )
        self.log.grid(row=1, column=0, sticky="nsew")

    def active_step(self) -> dict[str, object]:
        step = dict(self.steps[self.current_step])
        if self.current_step == 0 and self.vars["shipping_platform"].get() == "sf":
            step.update(
                {
                    "short": "顺丰上传",
                    "nav": "DP订单 -> 顺丰模板",
                    "title": "生成顺丰国际上传模板",
                    "summary": "从 DP 订单 CSV 生成顺丰国际批量导入 Excel。",
                    "config_note": "使用固定配置：SKU商品库、顺丰国际标准模板。请确认业务类型和是否带电。",
                    "button": "生成顺丰模板",
                    "after": "生成后到顺丰国际后台上传。创建订单成功后，下载“顺丰订单数据”，再做第2步。",
                    "command": self.run_sf,
                }
            )
        elif self.current_step == 1 and self.vars["shipping_platform"].get() == "sf":
            step.update(
                {
                    "nav": "顺丰订单 -> DP上传",
                    "title": "生成 DP 顺丰发货回填模板",
                    "summary": "把顺丰国际运单号整理成 DP 后台可上传的发货文件。",
                    "when": "顺丰国际订单已创建，并已下载“顺丰订单数据”Excel 后使用。",
                    "files": [
                        (
                            "本次顺丰订单数据 Excel",
                            "sf_orders",
                            [("Excel", "*.xls *.xlsx"), ("Excel 97-2003", "*.xls"), ("Excel 工作簿", "*.xlsx")],
                        )
                    ],
                    "config_note": "使用固定配置：DP发货模板。承运商固定为 SF INTERNATIONAL。",
                    "button": "生成顺丰 DP 回填",
                    "after": "生成后上传到 DropShipZone 后台。完成后下载顺丰合并面单 PDF，再做第3步。",
                    "command": self.run_dp,
                }
            )
        elif self.current_step == 2 and self.vars["shipping_platform"].get() == "sf":
            step.update(
                {
                    "nav": "顺丰面单PDF -> 寄方分组",
                    "title": "按寄方姓名拆分顺丰面单",
                    "summary": "读取每页面单中的寄方姓名，生成每个供应商对应的合并 PDF。",
                    "when": "顺丰国际合并面单 PDF 已下载后使用。",
                    "files": [("本次顺丰国际面单 PDF", "sf_label_pdf", [("PDF", "*.pdf")])],
                    "config_note": "顺丰模式不需要 DP 订单、顺丰订单信息或 SKU 商品库。",
                    "button": "按寄方姓名拆分",
                    "after": "每个寄方会得到一个合并 PDF；总目录保留分组汇总和校验报告。",
                    "command": self.run_labels,
                }
            )
        return step

    def select_shipping_platform(self, platform: str):
        if self.running:
            messagebox.showwarning("正在处理", "当前任务还没有结束，请等待处理完成后再切换平台。")
            return
        self.vars["shipping_platform"].set(platform)
        if self.nav_buttons:
            if platform == "sf":
                self.nav_buttons[0].configure(text="01  顺丰上传  DP订单 -> 顺丰模板")
                self.nav_buttons[1].configure(text="02  DP回填  顺丰订单 -> DP上传")
                self.nav_buttons[2].configure(text="03  面单分拆  顺丰PDF -> 寄方分组")
            else:
                self.nav_buttons[0].configure(text="01  云途上传  DP订单 -> 云途模板")
                self.nav_buttons[1].configure(text="02  DP回填  云途订单 -> DP上传")
                self.nav_buttons[2].configure(text="03  面单分拣  面单ZIP -> 供应商包")
        if self.current_step in (0, 1, 2):
            self.render_step()

    def select_step(self, index: int):
        self.current_step = index
        for idx, button in enumerate(self.nav_buttons):
            if idx == index:
                button.configure(
                    bg=self.COLORS["surface"],
                    fg=self.COLORS["accent_dark"],
                    activebackground="#ffffff",
                    activeforeground=self.COLORS["accent_dark"],
                )
            else:
                button.configure(
                    bg="#dfe8f3",
                    fg=self.COLORS["muted"],
                    activebackground="#e8eff7",
                    activeforeground=self.COLORS["ink"],
                )
        self.render_step()

    def render_step(self):
        for child in self.step_panel.winfo_children():
            child.destroy()
        self.action_buttons.clear()
        self.result_buttons.clear()

        step = self.active_step()
        self.step_panel.grid_columnconfigure(0, weight=1)
        for row in range(4):
            self.step_panel.grid_rowconfigure(row, weight=0)

        header = tk.Frame(self.step_panel, bg=self.COLORS["surface"])
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        tk.Label(
            header,
            text=f"STEP {step['index']}",
            bg=self.COLORS["accent_soft"],
            fg=self.COLORS["accent_dark"],
            font=("Segoe UI", 9, "bold"),
            width=8,
            padx=6,
            pady=5,
        ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 13))
        tk.Label(
            header,
            text=step["title"],
            bg=self.COLORS["surface"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 16, "bold"),
        ).grid(row=0, column=1, sticky="w")
        tk.Label(
            header,
            text=f"{step['summary']}  {step['when']}",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
            font=("Microsoft YaHei UI", 9),
            wraplength=820,
            justify="left",
        ).grid(row=1, column=1, sticky="w", pady=(3, 0))

        body = self.card(self.step_panel, pad=13, bg=self.COLORS["soft"])
        body.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        body.grid_columnconfigure(0, weight=1)

        body_head = tk.Frame(body, bg=self.COLORS["soft"])
        body_head.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        body_head.grid_columnconfigure(1, weight=1)
        tk.Label(
            body_head,
            text="本次需要提供的文件",
            bg=self.COLORS["soft"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 11, "bold"),
        ).grid(row=0, column=0, sticky="w")
        if self.current_step in (0, 1, 2):
            self.platform_selector(body_head).grid(row=0, column=1, sticky="w", padx=(18, 0))

        fixed_tip = tk.Frame(body_head, bg=self.COLORS["soft"])
        fixed_tip.grid(row=0, column=2, sticky="e", padx=(12, 0))
        sf_sender_split = self.current_step == 2 and self.vars["shipping_platform"].get() == "sf"
        tk.Label(
            fixed_tip,
            text="无需固定配置" if sf_sender_split else "固定配置",
            bg=self.COLORS["soft"],
            fg=self.COLORS["muted"],
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left")
        self.info_dot(
            fixed_tip,
            (
                "顺丰面单按每页的寄方姓名直接拆分，不使用订单表、SKU 商品库或固定模板。"
                if sf_sender_split
                else "SKU商品库、云途/顺丰标准模板、DP发货模板通常不用每天选择。只有模板或SKU信息更新时，才到右上角“固定配置”里维护。"
            ),
            bg=self.COLORS["soft"],
        ).pack(side="left", padx=(4, 0))

        for row, (label, key, filetypes) in enumerate(step["files"], start=1):
            self.file_field(body, row, label, key, filetypes)

        if self.current_step == 0 and self.vars["shipping_platform"].get() == "sf":
            self.sf_options_field(body, len(step["files"]) + 1)

        action = tk.Frame(
            self.step_panel,
            bg=self.COLORS["surface"],
            padx=12,
            pady=9,
            highlightthickness=1,
            highlightbackground=self.COLORS["line"],
        )
        action.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        action.grid_columnconfigure(0, weight=1)
        action_hint = tk.Frame(action, bg=self.COLORS["surface"])
        action_hint.grid(row=0, column=0, sticky="w")
        tk.Label(
            action_hint,
            text="确认文件后开始处理",
            bg=self.COLORS["surface"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 9, "bold"),
        ).pack(side="left")
        self.info_dot(
            action_hint,
            f"{step['config_note']}\n如果缺少必要文件，工具会停止并在日志里提示原因。",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
        ).pack(side="left", padx=(6, 0))
        tk.Frame(action_hint, bg=self.COLORS["line"], width=1, height=18).pack(side="left", padx=14)
        tk.Label(
            action_hint,
            text="完成后",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left")
        self.info_dot(
            action_hint,
            step["after"],
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
        ).pack(side="left", padx=(6, 0))
        start_button = self.primary_button(action, step["button"], step["command"])
        start_button.grid(row=0, column=1, sticky="e", padx=(12, 0))
        self.action_buttons.append(start_button)
        result_button = self.secondary_button(action, "打开结果文件夹", self.open_last_result_dir, compact=True)
        result_button.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.result_buttons.append(result_button)

        self.set_running(self.running)
        self.update_result_buttons()

    def platform_selector(self, parent):
        selector = tk.Frame(
            parent,
            bg="#dfe8f3",
            padx=3,
            pady=3,
            highlightthickness=1,
            highlightbackground="#d4dfec",
        )
        self.platform_buttons.clear()
        current = self.vars["shipping_platform"].get()
        for column, (key, label) in enumerate((("yunexpress", "云途"), ("sf", "顺丰国际"))):
            selected = key == current
            button = tk.Button(
                selector,
                text=label,
                command=lambda value=key: self.select_shipping_platform(value),
                bg=self.COLORS["surface"] if selected else "#dfe8f3",
                fg=self.COLORS["accent_dark"] if selected else self.COLORS["muted"],
                activebackground="#ffffff",
                activeforeground=self.COLORS["accent_dark"],
                relief="flat",
                bd=0,
                padx=12,
                pady=3,
                cursor="hand2",
                font=("Microsoft YaHei UI", 9, "bold"),
            )
            button.grid(row=0, column=column, padx=(0, 3) if column == 0 else 0)
            self.platform_buttons[key] = button
        return selector

    def sf_options_field(self, parent, row):
        field = tk.Frame(parent, bg=self.COLORS["soft"])
        field.grid(row=row, column=0, sticky="ew", pady=(5, 2))
        field.grid_columnconfigure(1, weight=1)

        tk.Label(
            field,
            text="顺丰业务类型",
            bg=self.COLORS["soft"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 9, "bold"),
            width=25,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        business_type = ttk.Combobox(
            field,
            textvariable=self.vars["sf_business_type"],
            values=generate_sf_international_template.SF_BUSINESS_TYPES,
            state="readonly",
            font=("Microsoft YaHei UI", 9),
        )
        business_type.grid(row=0, column=1, sticky="ew", padx=(0, 12))

        tk.Label(
            field,
            text="是否带电",
            bg=self.COLORS["soft"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 9, "bold"),
        ).grid(row=0, column=2, sticky="e", padx=(0, 8))
        battery = ttk.Combobox(
            field,
            textvariable=self.vars["sf_battery"],
            values=generate_sf_international_template.BATTERY_OPTIONS,
            state="readonly",
            width=7,
            font=("Microsoft YaHei UI", 9),
        )
        battery.grid(row=0, column=3, sticky="e")

    def file_field(self, parent, row, label, key, filetypes):
        field = tk.Frame(parent, bg=self.COLORS["soft"])
        field.grid(row=row, column=0, sticky="ew", pady=3)
        field.grid_columnconfigure(1, weight=1)
        tk.Label(
            field,
            text=label,
            bg=self.COLORS["soft"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 9, "bold"),
            width=25,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.entry(field, self.vars[key]).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.secondary_button(field, "选择", lambda: self.pick_file(key, filetypes), compact=True).grid(
            row=0, column=2, sticky="e"
        )

    def card(self, parent, pad=12, bg=None):
        color = bg or self.COLORS["surface"]
        return tk.Frame(
            parent,
            bg=color,
            padx=pad,
            pady=pad,
            highlightthickness=1,
            highlightbackground=self.COLORS["line"],
        )

    def entry(self, parent, variable, width=None, state="normal"):
        return tk.Entry(
            parent,
            textvariable=variable,
            width=width,
            state=state,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="#cad7e6",
            highlightcolor=self.COLORS["accent"],
            bg="#ffffff",
            fg=self.COLORS["ink"],
            readonlybackground="#edf3f9",
            font=("Microsoft YaHei UI", 9),
        )

    def primary_button(self, parent, text, command, inverted=False):
        bg = "#ffffff" if inverted else self.COLORS["accent"]
        fg = self.COLORS["accent_dark"] if inverted else "#ffffff"
        active_bg = "#edf6ff" if inverted else self.COLORS["accent_dark"]
        active_fg = self.COLORS["accent_dark"] if inverted else "#ffffff"
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=active_fg,
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            font=("Microsoft YaHei UI", 10, "bold"),
        )

    def secondary_button(self, parent, text, command, compact=False):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=self.COLORS["button_soft"],
            fg=self.COLORS["ink"],
            activebackground="#dfeaf6",
            activeforeground=self.COLORS["ink"],
            relief="flat",
            bd=0,
            padx=8 if compact else 11,
            pady=3 if compact else 6,
            cursor="hand2",
            font=("Microsoft YaHei UI", 9),
        )

    def info_dot(self, parent, text, bg=None, fg=None):
        color = bg or self.COLORS["surface"]
        outline = fg or "#74839a"
        hover_fill = "#e4f1ff"
        mark = tk.Canvas(
            parent,
            width=17,
            height=17,
            bg=color,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        circle = mark.create_oval(2, 2, 15, 15, outline=outline, width=1)
        mark.create_text(8.5, 8.0, text="i", fill=outline, font=("Segoe UI", 8, "bold"))

        def on_enter(_event=None):
            mark.itemconfigure(circle, fill=hover_fill)

        def on_leave(_event=None):
            mark.itemconfigure(circle, fill=color)

        mark.bind("<Enter>", on_enter, add="+")
        mark.bind("<Leave>", on_leave, add="+")
        Tooltip(mark, text)
        return mark

    def pick_file(self, key, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes + [("所有文件", "*.*")])
        if path:
            self.vars[key].set(path)

    def open_config_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("固定配置")
        dialog.geometry("820x410")
        dialog.minsize(760, 390)
        dialog.configure(bg=self.COLORS["bg"])
        dialog.transient(self.root)
        dialog.grab_set()

        panel = self.card(dialog, pad=16)
        panel.pack(fill="both", expand=True, padx=16, pady=16)
        panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            panel,
            text="固定配置",
            bg=self.COLORS["surface"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            panel,
            text="这些文件通常不需要每天修改。只有 SKU 商品库或平台模板更新时，才需要在这里重新选择。",
            bg=self.COLORS["surface"],
            fg=self.COLORS["muted"],
            font=("Microsoft YaHei UI", 9),
            wraplength=760,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(5, 12))

        for row, (label, key, filetypes) in enumerate(self.config_files, start=2):
            self.config_file_field(panel, row, label, key, filetypes)

        actions = tk.Frame(panel, bg=self.COLORS["surface"])
        actions.grid(row=len(self.config_files) + 2, column=0, sticky="ew", pady=(14, 0))
        actions.grid_columnconfigure(0, weight=1)
        self.secondary_button(actions, "自动识别固定配置", self.autofill_config).grid(row=0, column=0, sticky="w")
        self.primary_button(actions, "保存并关闭", dialog.destroy).grid(row=0, column=1, sticky="e")

    def config_file_field(self, parent, row, label, key, filetypes):
        field = tk.Frame(parent, bg=self.COLORS["surface"])
        field.grid(row=row, column=0, sticky="ew", pady=6)
        field.grid_columnconfigure(1, weight=1)
        tk.Label(
            field,
            text=label,
            bg=self.COLORS["surface"],
            fg=self.COLORS["ink"],
            font=("Microsoft YaHei UI", 9, "bold"),
            width=16,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.entry(field, self.vars[key]).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.secondary_button(field, "选择", lambda: self.pick_file(key, filetypes), compact=True).grid(
            row=0, column=2, sticky="e"
        )

    def refresh_output_dir(self):
        try:
            self.output_dir.set(str(date_folder(self.vars["date"].get())))
            return True
        except Exception:
            messagebox.showerror("日期格式错误", "请使用 YYYY-MM-DD，例如 2026-07-10")
            return False

    def autofill_latest(self):
        self.vars["dp_orders"].set(latest_file(DEFAULT_DOWNLOADS, "Exported_Orders_*.csv"))
        self.vars["yun_orders"].set(latest_file(DEFAULT_DOWNLOADS, "订单信息_*.xlsx"))
        self.vars["sf_orders"].set(latest_file(DEFAULT_DOWNLOADS, "顺丰订单数据*.xls*"))
        self.vars["label_zip"].set(latest_file(DEFAULT_DOWNLOADS, "运单标签*.zip"))
        self.vars["sf_label_pdf"].set(latest_file(DEFAULT_DOWNLOADS, "labelSF*.pdf"))
        self.autofill_config(write_log=False)
        self.write("已重新识别本次批次文件，并同步检查固定配置。请确认路径是否对应当前批次。")

    def autofill_config(self, write_log=True):
        self.vars["sku"].set(latest_file(DEFAULT_SKU_DIR, "*SKU*.xlsx"))
        self.vars["yun_template"].set(latest_file(DEFAULT_TEMPLATE_DIR, "*CN0C021578*.xlsx"))
        self.vars["sf_template"].set(
            latest_file_across(
                [
                    (DEFAULT_DOWNLOADS, "BatchImportOrders_*.xlsm"),
                    (DEFAULT_TEMPLATE_DIR, "BatchImportOrders_*.xlsm"),
                ]
            )
        )
        self.vars["dp_template"].set(latest_file(DEFAULT_TEMPLATE_DIR, "*Shipment_sample.xlsx"))
        if write_log:
            self.write("已重新识别固定配置。SKU 商品库和模板如果没有更新，平时不需要改。")

    def open_output_dir(self):
        if not self.refresh_output_dir():
            return
        path = Path(self.output_dir.get())
        path.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer", str(path)])

    def open_last_result_dir(self):
        if not self.last_result_dir or not self.last_result_dir.exists():
            messagebox.showwarning("暂无结果", "当前还没有可打开的成功结果文件夹。")
            return
        subprocess.Popen(["explorer", str(self.last_result_dir)])

    def extract_result_dir(self, output: str) -> Path | None:
        matches = re.findall(r"^OK:\s*(.+?)\s*$", output, flags=re.MULTILINE)
        if not matches:
            return None
        result_path = Path(matches[-1].strip().strip('"'))
        if result_path.is_file():
            return result_path.parent
        if result_path.is_dir():
            return result_path
        return result_path.parent

    def remember_result_dir(self, result_dir: Path | None):
        self.last_result_dir = result_dir
        self.update_result_buttons()
        if result_dir:
            self.status_text.set("成功，可打开结果文件夹")

    def update_result_buttons(self):
        enabled = bool(self.last_result_dir and self.last_result_dir.exists())
        for button in self.result_buttons:
            button.configure(state="normal" if enabled else "disabled")

    def write(self, text):
        if threading.current_thread() is threading.main_thread():
            self._append_log(text)
        else:
            self.root.after(0, self._append_log, text)

    def _append_log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def require_files(self, keys):
        missing = [
            self.FILE_LABELS.get(key, key)
            for key in keys
            if not self.vars[key].get() or not Path(self.vars[key].get()).exists()
        ]
        if missing:
            messagebox.showerror(
                "文件缺失",
                "这些文件没有找到。若是前面步骤已选择过，请点击顶部“自动识别”或回到对应步骤检查：\n"
                + "\n".join(missing),
            )
            return False
        return True

    def set_running(self, running: bool):
        self.running = running
        if running:
            self.status_text.set("处理中，请等待")
        elif self.last_result_dir and self.last_result_dir.exists():
            self.status_text.set("成功，可打开结果文件夹")
        else:
            self.status_text.set("就绪")
        for button in self.action_buttons:
            button.configure(state="disabled" if running else "normal")
        if running:
            for button in self.result_buttons:
                button.configure(state="disabled")
        else:
            self.update_result_buttons()

    def run_script(self, script_name: str, args: list[str]) -> tuple[int, str, str]:
        module = SCRIPT_MODULES[script_name]
        old_argv = sys.argv[:]
        stdout = io.StringIO()
        stderr = io.StringIO()
        try:
            sys.argv = [script_name] + args
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                code = module.main()
        finally:
            sys.argv = old_argv
        return int(code or 0), stdout.getvalue(), stderr.getvalue()

    def run_cmd(self, title, cmd):
        if self.running:
            messagebox.showwarning("正在处理", "当前任务还没有结束，请等待运行日志显示结果。")
            return
        self.last_result_dir = None
        self.update_result_buttons()
        self.set_running(True)

        def worker():
            self.write("")
            self.write(f"开始：{title}")
            returncode, stdout, stderr = self.run_script(cmd[0], cmd[1:])
            if stdout.strip():
                self.write(stdout.strip())
            if stderr.strip():
                self.write(stderr.strip())
            if returncode == 0:
                self.write("结果：成功")
                result_dir = self.extract_result_dir(stdout)
                self.root.after(0, lambda: self.remember_result_dir(result_dir))
                if result_dir:
                    self.write(f"结果文件夹：{result_dir}")
            else:
                self.write(f"结果：失败，退出码 {returncode}")
            self.root.after(0, lambda: self.set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def base_cmd(self, script):
        if not self.refresh_output_dir():
            raise ValueError("invalid date")
        return [script, "--date", self.vars["date"].get(), "--output-root", self.output_dir.get()]

    def run_sf(self):
        if not self.require_files(["dp_orders", "sku", "sf_template"]):
            return
        business_type = self.vars["sf_business_type"].get()
        if business_type not in generate_sf_international_template.SF_BUSINESS_TYPES:
            messagebox.showerror("请选择业务类型", "请先选择本批顺丰国际订单使用的业务类型。")
            return
        self.run_cmd(
            "第1步 生成顺丰国际上传模板",
            self.base_cmd("generate_sf_international_template.py")
            + [
                "--dp-orders-csv",
                self.vars["dp_orders"].get(),
                "--sku-xlsx",
                self.vars["sku"].get(),
                "--sf-template-xlsm",
                self.vars["sf_template"].get(),
                "--business-type",
                business_type,
                "--battery",
                self.vars["sf_battery"].get(),
            ],
        )

    def run_yun(self):
        if not self.require_files(["dp_orders", "sku", "yun_template"]):
            return
        self.run_cmd(
            "第1步 生成云途上传模板",
            self.base_cmd("generate_yunexpress_template.py")
            + [
                "--dp-orders-csv",
                self.vars["dp_orders"].get(),
                "--sku-xlsx",
                self.vars["sku"].get(),
                "--yunexpress-template-xlsx",
                self.vars["yun_template"].get(),
            ],
        )

    def run_dp(self):
        platform = self.vars["shipping_platform"].get()
        source_key = "sf_orders" if platform == "sf" else "yun_orders"
        if not self.require_files([source_key, "dp_template"]):
            return
        source_args = (
            ["--sf-international-file", self.vars["sf_orders"].get()]
            if platform == "sf"
            else ["--yunexpress-xlsx", self.vars["yun_orders"].get()]
        )
        title = "第2步 生成顺丰 DP 回填模板" if platform == "sf" else "第2步 生成云途 DP 回填模板"
        self.run_cmd(
            title,
            self.base_cmd("generate_dp_shipment_upload.py")
            + source_args
            + [
                "--shipment-template-xlsx",
                self.vars["dp_template"].get(),
            ],
        )

    def run_labels(self):
        if self.vars["shipping_platform"].get() == "sf":
            if not self.require_files(["sf_label_pdf"]):
                return
            self.run_cmd(
                "第3步 按寄方姓名拆分顺丰面单",
                self.base_cmd("split_sf_labels_by_sender.py")
                + ["--pdf", self.vars["sf_label_pdf"].get()],
            )
            return
        if not self.require_files(["label_zip", "yun_orders", "dp_orders", "sku"]):
            return
        self.run_cmd(
            "第3步 按 SKU 分拣面单",
            self.base_cmd("sort_yunexpress_labels_by_sku.py")
            + [
                "--zip",
                self.vars["label_zip"].get(),
                "--yunexpress-xlsx",
                self.vars["yun_orders"].get(),
                "--dp-orders-csv",
                self.vars["dp_orders"].get(),
                "--sku-xlsx",
                self.vars["sku"].get(),
            ],
        )


if __name__ == "__main__":
    app_root = tk.Tk()
    App(app_root)
    app_root.mainloop()
