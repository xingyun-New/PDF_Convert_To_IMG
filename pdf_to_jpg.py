"""
PDFtoJPG - GUI 版
=================
Tkinter 图形界面: 用户选择输入/输出目录, 设置格式/DPI/质量, 可选去除页码,
后台线程批量把 PDF 逐页渲染输出到 <输出目录>/<PDF文件名>/page_NNN.{jpg,png}

设计要点:
- PyMuPDF (fitz) 渲染, 无外部二进制依赖, 适合 PyInstaller --onefile + --windowed 打包。
- 转换在后台 worker 线程, 通过 queue.Queue 把进度/日志推回 UI 线程, 主线程 50ms 轮询。
- 页码抹除分两层: 基于文本的 PyMuPDF redact 智能识别 + 渲染时直接 clip 裁掉页脚/页眉。
- 设置持久化到 exe 同目录的 .pdftojpg.ini, 不存在或损坏时静默回退默认。
"""

from __future__ import annotations

import configparser
import os
import queue
import re
import sys
import threading
import time
import traceback
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


# ---------- 控制台 UTF-8 (脚本调试时有用; --windowed exe 下无 stdout, 此处是 no-op) ----------

def _setup_console_utf8() -> None:
    if os.name == "nt":
        try:
            os.system("chcp 65001 > nul")
        except Exception:
            pass
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None:
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


_setup_console_utf8()


try:
    import fitz  # PyMuPDF
except ImportError:
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "缺少依赖",
        "未找到 PyMuPDF (fitz)。\n\n请在命令行运行:\n    py -m pip install -r requirements.txt",
    )
    sys.exit(1)


# ---------- 路径定位 ----------

def app_dir() -> Path:
    """返回 exe / 脚本所在目录 (打包后用 sys.executable 的目录, 而非 _MEIPASS)。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------- 页码抹除 ----------

PAGE_NUMBER_PATTERNS = [
    re.compile(r"^\s*\d{1,4}\s*$"),                                  # 1
    re.compile(r"^\s*\d{1,4}\s*/\s*\d{1,4}\s*$"),                    # 1/12
    re.compile(r"^\s*第\s*\d{1,4}\s*页\s*$"),                        # 第 1 页
    re.compile(r"^\s*[Pp]age\s+\d{1,4}(\s+of\s+\d{1,4})?\s*$"),      # Page 1 of 12
    re.compile(r"^\s*[-—]\s*\d{1,4}\s*[-—]\s*$"),                    # - 1 -
    re.compile(r"^\s*\d{1,4}\s*[\.,]\s*$"),                          # 1.
]


@dataclass
class RedactOptions:
    """页码处理三个独立项, 三者可任意组合, 互不干扰。

    - smart: 是否启用文本级智能识别并抹白页码 (PyMuPDF redact)
    - crop_top_pct: 渲染时从页面顶部裁剪的高度比例 (0 表示不裁)
    - crop_bottom_pct: 渲染时从页面底部裁剪的高度比例 (0 表示不裁)
    - scan_band: 智能识别只在页面顶/底各 N% 区域内查找数字, 防止误擦正文
    - enabled: 总开关, 关闭时三项全部失效, 等价于不抹除/不裁剪
    """
    enabled: bool = True
    smart: bool = True
    crop_top_pct: float = 0.0
    crop_bottom_pct: float = 0.05
    scan_band: float = 0.12


def _redact_page_numbers(page, band: float = 0.12) -> int:
    """在 page 对象上抹除页码文字, 返回抹除次数。仅修改内存对象, 不写回原文件。"""
    pr = page.rect
    top_y = pr.y0 + pr.height * band
    bot_y = pr.y1 - pr.height * band
    n = 0
    try:
        text_dict = page.get_text("dict")
    except Exception:
        return 0
    for blk in text_dict.get("blocks", []):
        if blk.get("type") != 0:  # 0=text, 1=image
            continue
        for line in blk.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text or len(text) > 24:
                    continue
                if not any(p.match(text) for p in PAGE_NUMBER_PATTERNS):
                    continue
                bbox = fitz.Rect(span["bbox"])
                # 只抹除位于顶/底 band 区域内的命中, 防止误擦正文里的数字
                if bbox.y1 <= top_y or bbox.y0 >= bot_y:
                    try:
                        page.add_redact_annot(bbox, fill=(1, 1, 1))
                        n += 1
                    except Exception:
                        pass
    if n:
        try:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        except Exception:
            try:
                page.apply_redactions()
            except Exception:
                return 0
    return n


def _crop_clip(page, redact: Optional[RedactOptions]):
    """根据 redact 配置返回渲染时的 clip 矩形 (同时支持页眉与页脚独立裁剪)。"""
    if not (redact and redact.enabled):
        return None
    top = max(0.0, min(0.45, float(redact.crop_top_pct)))
    bot = max(0.0, min(0.45, float(redact.crop_bottom_pct)))
    if top <= 0.0 and bot <= 0.0:
        return None
    pr = page.rect
    x0, y0, x1, y1 = pr.x0, pr.y0, pr.x1, pr.y1
    if top > 0.0:
        y0 += pr.height * top
    if bot > 0.0:
        y1 -= pr.height * bot
    if y1 - y0 <= 1.0:
        return None
    return fitz.Rect(x0, y0, x1, y1)


# ---------- 扫描与转换 ----------

def find_pdfs(input_dir: Path) -> List[Path]:
    """递归收集 input_dir 下所有 .pdf, 按路径排序。"""
    pdfs = [p for p in input_dir.rglob("*.pdf") if p.is_file()]
    pdfs.sort(key=lambda p: str(p).lower())
    return pdfs


def safe_subdir_name(pdf_path: Path, input_dir: Path) -> Path:
    """构造输出子目录的相对路径 (保留 input/ 下的层级, 去掉 .pdf 扩展名)。"""
    try:
        rel = pdf_path.relative_to(input_dir)
    except ValueError:
        rel = Path(pdf_path.name)
    return rel.with_suffix("")


def get_page_count(pdf_path: Path) -> int:
    try:
        with fitz.open(pdf_path) as doc:
            return doc.page_count
    except Exception:
        return 0


@dataclass
class FileResult:
    pdf: Path
    ok: int = 0
    fail: int = 0
    redacted: int = 0
    errors: List[str] = field(default_factory=list)


PageCallback = Callable[[Path, int, int, int], None]


def convert_one_pdf(
    pdf_path: Path,
    input_dir: Path,
    output_dir: Path,
    dpi: int,
    fmt: str,
    quality: int,
    redact: Optional[RedactOptions] = None,
    progress_cb: Optional[PageCallback] = None,
    cancel_event: Optional[threading.Event] = None,
) -> FileResult:
    """
    转换单个 PDF。progress_cb(pdf, page_no, total, redacted_n) 在每页完成后调用,
    cancel_event 在每页边界检查, 命中后立即返回当前 result。
    """
    result = FileResult(pdf=pdf_path)
    sub = output_dir / safe_subdir_name(pdf_path, input_dir)
    try:
        sub.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        result.errors.append(f"创建输出目录失败: {e}")
        return result
    ext = "jpg" if fmt == "jpeg" else "png"

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        result.errors.append(f"无法打开 PDF: {e}")
        return result

    try:
        total = doc.page_count
        if total == 0:
            result.errors.append("PDF 为空 (0 页)")
            return result

        for i in range(total):
            if cancel_event is not None and cancel_event.is_set():
                return result
            page_no = i + 1
            try:
                page = doc.load_page(i)
                redacted_n = 0
                if redact and redact.enabled and redact.smart:
                    redacted_n = _redact_page_numbers(page, band=redact.scan_band)
                    result.redacted += redacted_n
                clip = _crop_clip(page, redact)
                pix = page.get_pixmap(dpi=dpi, clip=clip) if clip is not None else page.get_pixmap(dpi=dpi)
                out_file = sub / f"page_{page_no:03d}.{ext}"
                if fmt == "jpeg":
                    pix.save(str(out_file), output="jpeg", jpg_quality=quality)
                else:
                    pix.save(str(out_file), output="png")
                result.ok += 1
                if progress_cb is not None:
                    progress_cb(pdf_path, page_no, total, redacted_n)
            except Exception as e:
                result.fail += 1
                result.errors.append(f"第 {page_no} 页: {e}")
    finally:
        doc.close()

    return result


# ---------- 设置持久化 ----------

SETTINGS_FILENAME = ".pdftojpg.ini"


@dataclass
class Settings:
    input_dir: str = ""
    output_dir: str = ""
    fmt: str = "jpeg"
    dpi: int = 200
    quality: int = 95
    redact_enabled: bool = True
    redact_smart: bool = True              # 智能识别页码 (独立开关)
    crop_top_pct: int = 0                  # 页眉裁剪百分比 0-30
    crop_bottom_pct: int = 5               # 页脚裁剪百分比 0-30
    scan_band: int = 12                    # 智能扫描带宽 5-30

    @classmethod
    def load(cls, path: Path) -> "Settings":
        s = cls()
        if not path.exists():
            return s
        try:
            cp = configparser.ConfigParser()
            cp.read(path, encoding="utf-8")
            if "general" in cp:
                g = cp["general"]
                s.input_dir = g.get("input_dir", s.input_dir)
                s.output_dir = g.get("output_dir", s.output_dir)
                s.fmt = g.get("fmt", s.fmt)
                s.dpi = g.getint("dpi", s.dpi)
                s.quality = g.getint("quality", s.quality)
                s.redact_enabled = g.getboolean("redact_enabled", s.redact_enabled)
                s.scan_band = g.getint("scan_band", s.scan_band)
                # 优先读新字段
                if "redact_smart" in g or "crop_top_pct" in g or "crop_bottom_pct" in g:
                    s.redact_smart = g.getboolean("redact_smart", s.redact_smart)
                    s.crop_top_pct = g.getint("crop_top_pct", s.crop_top_pct)
                    s.crop_bottom_pct = g.getint("crop_bottom_pct", s.crop_bottom_pct)
                else:
                    # 旧版 ini 迁移: redact_mode + crop_pct -> smart + top/bottom
                    old_mode = g.get("redact_mode", "smart_plus_crop")
                    old_pct = g.getint("crop_pct", 5)
                    s.redact_smart = old_mode in ("smart", "smart_plus_crop")
                    s.crop_top_pct = old_pct if old_mode == "crop_top" else 0
                    s.crop_bottom_pct = old_pct if old_mode in ("crop_bottom", "smart_plus_crop") else 0
        except Exception:
            pass
        return s

    def save(self, path: Path) -> None:
        try:
            cp = configparser.ConfigParser()
            cp["general"] = {k: str(v) for k, v in asdict(self).items()}
            with path.open("w", encoding="utf-8") as f:
                cp.write(f)
        except Exception:
            pass


# ---------- GUI ----------

FMT_OPTIONS = [("JPG (推荐)", "jpeg"), ("PNG (无损)", "png")]
DPI_OPTIONS = [100, 150, 200, 300, 450, 600]
QUALITY_OPTIONS = [80, 85, 90, 95, 100]


def _open_in_explorer(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(path))
        else:
            os.system(f'xdg-open "{path}"')
    except Exception as e:
        messagebox.showerror("无法打开目录", str(e))


class App(tk.Tk):
    def __init__(self, settings: Settings, settings_path: Path):
        super().__init__()
        self.settings_path = settings_path

        self.title("PDFtoJPG 转换器")
        self.geometry("860x740")
        self.minsize(760, 640)

        # 中文字体: 优先用微软雅黑, 没有则交给 Tk 默认
        try:
            self.option_add("*Font", "{Microsoft YaHei UI} 10")
        except Exception:
            pass

        self.q: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: Optional[threading.Thread] = None
        self.scan_pdfs: List[Tuple[Path, int]] = []   # (pdf_path, page_count)
        self.total_pages_all: int = 0
        self.overall_done: int = 0
        self.start_time: float = 0.0

        # tk 变量
        self.var_input = tk.StringVar()
        self.var_output = tk.StringVar()
        self.var_fmt = tk.StringVar(value="jpeg")
        self.var_dpi = tk.IntVar(value=200)
        self.var_quality = tk.IntVar(value=95)
        self.var_redact_enabled = tk.BooleanVar(value=True)
        self.var_redact_smart = tk.BooleanVar(value=True)
        self.var_crop_top_pct = tk.IntVar(value=0)
        self.var_crop_bottom_pct = tk.IntVar(value=5)
        self.var_scan_band = tk.IntVar(value=12)
        self.var_progress = tk.DoubleVar(value=0.0)
        self.var_progress_text = tk.StringVar(value="0.0%")
        self.var_status = tk.StringVar(value="就绪")

        self._build_widgets()
        self._apply_settings(settings)
        self._update_quality_state()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(50, self._poll_queue)

    # ----- UI 构建 -----

    def _build_widgets(self) -> None:
        pad = {"padx": 8, "pady": 4}
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=10)
        main.columnconfigure(1, weight=1)

        # 输入/输出目录
        ttk.Label(main, text="输入文件夹:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(main, textvariable=self.var_input).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(main, text="浏览...", command=self._browse_input).grid(row=0, column=2, sticky="ew", **pad)

        ttk.Label(main, text="输出文件夹:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(main, textvariable=self.var_output).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(main, text="浏览...", command=self._browse_output).grid(row=1, column=2, sticky="ew", **pad)

        # 格式/DPI/质量
        params = ttk.Frame(main)
        params.grid(row=2, column=0, columnspan=3, sticky="ew", padx=4, pady=(8, 4))
        ttk.Label(params, text="格式:").pack(side="left", padx=(4, 4))
        cb_fmt = ttk.Combobox(params, width=12, state="readonly",
                              values=[o[0] for o in FMT_OPTIONS])
        cb_fmt.current(0)
        cb_fmt.pack(side="left", padx=(0, 12))
        cb_fmt.bind("<<ComboboxSelected>>", lambda e: self._on_fmt_pick(cb_fmt))
        self._cb_fmt = cb_fmt

        ttk.Label(params, text="DPI:").pack(side="left", padx=(0, 4))
        cb_dpi = ttk.Combobox(params, width=8, state="readonly",
                              values=[str(d) for d in DPI_OPTIONS])
        cb_dpi.set("200")
        cb_dpi.pack(side="left", padx=(0, 12))
        cb_dpi.bind("<<ComboboxSelected>>", lambda e: self.var_dpi.set(int(cb_dpi.get())))
        self._cb_dpi = cb_dpi

        ttk.Label(params, text="JPG 质量:").pack(side="left", padx=(0, 4))
        cb_q = ttk.Combobox(params, width=6, state="readonly",
                            values=[str(q) for q in QUALITY_OPTIONS])
        cb_q.set("95")
        cb_q.pack(side="left", padx=(0, 12))
        cb_q.bind("<<ComboboxSelected>>", lambda e: self.var_quality.set(int(cb_q.get())))
        self._cb_q = cb_q

        # 页码处理分组 - 三项独立: 智能识别 / 页眉裁剪 / 页脚裁剪 (可任意组合)
        rg = ttk.LabelFrame(main, text="页码处理")
        rg.grid(row=3, column=0, columnspan=3, sticky="ew", padx=4, pady=6)
        rg.columnconfigure(7, weight=1)

        ttk.Checkbutton(
            rg, text="去除每页的页码 (总开关)", variable=self.var_redact_enabled,
            command=self._update_redact_state,
        ).grid(row=0, column=0, columnspan=8, sticky="w", padx=8, pady=(6, 2))

        self._cb_smart = ttk.Checkbutton(
            rg, text="智能识别页码并抹除", variable=self.var_redact_smart,
            command=self._update_redact_state,
        )
        self._cb_smart.grid(row=1, column=0, columnspan=4, sticky="w", padx=(8, 4), pady=2)

        ttk.Label(rg, text="智能扫描带宽:").grid(row=1, column=4, sticky="e", padx=(16, 4), pady=2)
        self._sp_band = ttk.Spinbox(rg, from_=5, to=30, width=5, textvariable=self.var_scan_band)
        self._sp_band.grid(row=1, column=5, sticky="w", padx=4, pady=2)
        ttk.Label(rg, text="%").grid(row=1, column=6, sticky="w", pady=2)

        ttk.Label(rg, text="页眉裁剪:").grid(row=2, column=0, sticky="e", padx=(8, 4), pady=(2, 8))
        self._sp_top = ttk.Spinbox(rg, from_=0, to=30, width=5, textvariable=self.var_crop_top_pct)
        self._sp_top.grid(row=2, column=1, sticky="w", padx=4, pady=(2, 8))
        ttk.Label(rg, text="%").grid(row=2, column=2, sticky="w", pady=(2, 8))
        ttk.Label(rg, text="(0 = 不裁顶部)").grid(row=2, column=3, sticky="w", padx=(2, 12), pady=(2, 8))

        ttk.Label(rg, text="页脚裁剪:").grid(row=2, column=4, sticky="e", padx=(8, 4), pady=(2, 8))
        self._sp_bot = ttk.Spinbox(rg, from_=0, to=30, width=5, textvariable=self.var_crop_bottom_pct)
        self._sp_bot.grid(row=2, column=5, sticky="w", padx=4, pady=(2, 8))
        ttk.Label(rg, text="%").grid(row=2, column=6, sticky="w", pady=(2, 8))
        ttk.Label(rg, text="(0 = 不裁底部)").grid(row=2, column=7, sticky="w", padx=(2, 8), pady=(2, 8))

        # PDF 列表
        list_frame = ttk.LabelFrame(main, text="待处理 PDF")
        list_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=4, pady=4)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        cols = ("name", "pages")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=6)
        self.tree.heading("name", text="文件 (相对输入目录)")
        self.tree.heading("pages", text="页数")
        self.tree.column("name", anchor="w", stretch=True)
        self.tree.column("pages", anchor="e", width=90, stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns", pady=4)

        # 按钮区
        btn = ttk.Frame(main)
        btn.grid(row=5, column=0, columnspan=3, sticky="ew", padx=4, pady=4)
        self.btn_scan = ttk.Button(btn, text="扫描", command=self._on_scan)
        self.btn_scan.pack(side="left", padx=4)
        self.btn_start = ttk.Button(btn, text="开始转换", command=self._on_start)
        self.btn_start.pack(side="left", padx=4)
        self.btn_cancel = ttk.Button(btn, text="取消", command=self._on_cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=4)
        self.btn_open = ttk.Button(btn, text="打开输出目录", command=self._on_open_output)
        self.btn_open.pack(side="left", padx=4)

        # 进度条 + 百分比 + 状态
        prog = ttk.Frame(main)
        prog.grid(row=6, column=0, columnspan=3, sticky="ew", padx=4, pady=(8, 2))
        prog.columnconfigure(0, weight=1)
        self.pb = ttk.Progressbar(prog, orient="horizontal", mode="determinate",
                                  variable=self.var_progress, maximum=100.0)
        self.pb.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.lbl_pct = ttk.Label(prog, textvariable=self.var_progress_text,
                                 width=8, anchor="e",
                                 font=("Microsoft YaHei UI", 11, "bold"))
        self.lbl_pct.grid(row=0, column=1, sticky="e")
        self.lbl_status = ttk.Label(main, textvariable=self.var_status)
        self.lbl_status.grid(row=7, column=0, columnspan=3, sticky="w", padx=4)

        # 日志
        log_frame = ttk.LabelFrame(main, text="日志")
        log_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=4, pady=6)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main.rowconfigure(8, weight=1)
        self.log = ScrolledText(log_frame, height=10, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.log.configure(state="disabled")

    # ----- 事件 -----

    def _browse_input(self) -> None:
        d = filedialog.askdirectory(initialdir=self.var_input.get() or str(app_dir()),
                                    title="选择输入文件夹 (PDF 所在)")
        if d:
            self.var_input.set(d)

    def _browse_output(self) -> None:
        d = filedialog.askdirectory(initialdir=self.var_output.get() or str(app_dir()),
                                    title="选择输出文件夹 (图片保存到这里)")
        if d:
            self.var_output.set(d)

    def _on_fmt_pick(self, cb: ttk.Combobox) -> None:
        idx = cb.current()
        if 0 <= idx < len(FMT_OPTIONS):
            self.var_fmt.set(FMT_OPTIONS[idx][1])
        self._update_quality_state()

    def _update_quality_state(self) -> None:
        state = "readonly" if self.var_fmt.get() == "jpeg" else "disabled"
        try:
            self._cb_q.configure(state=state)
        except Exception:
            pass

    def _update_redact_state(self) -> None:
        en = self.var_redact_enabled.get()
        smart = en and self.var_redact_smart.get()
        self._cb_smart.configure(state=("normal" if en else "disabled"))
        self._sp_top.configure(state=("normal" if en else "disabled"))
        self._sp_bot.configure(state=("normal" if en else "disabled"))
        # 扫描带宽仅在「智能识别」实际启用时有意义
        self._sp_band.configure(state=("normal" if smart else "disabled"))

    def _on_scan(self) -> None:
        input_dir = Path(self.var_input.get()).expanduser()
        if not input_dir.exists() or not input_dir.is_dir():
            messagebox.showwarning("输入目录无效", f"目录不存在或不可访问:\n{input_dir}")
            return
        for it in self.tree.get_children():
            self.tree.delete(it)
        self.scan_pdfs = []
        pdfs = find_pdfs(input_dir)
        if not pdfs:
            self._log(f"在 {input_dir} 下未找到 PDF 文件。")
            self.var_status.set(f"未找到 PDF (扫描目录: {input_dir})")
            return
        for p in pdfs:
            pc = get_page_count(p)
            try:
                rel = p.relative_to(input_dir)
            except ValueError:
                rel = p
            self.tree.insert("", "end", values=(str(rel), pc))
            self.scan_pdfs.append((p, pc))
        total_pages = sum(pc for _, pc in self.scan_pdfs)
        self.total_pages_all = total_pages
        self.var_status.set(f"扫描完成: {len(pdfs)} 个 PDF, 共 {total_pages} 页")
        self._log(f"扫描完成: {len(pdfs)} 个 PDF, 共 {total_pages} 页")

    def _on_start(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            return
        if not self.scan_pdfs:
            self._on_scan()
            if not self.scan_pdfs:
                return

        input_dir = Path(self.var_input.get()).expanduser()
        output_dir = Path(self.var_output.get()).expanduser()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("无法创建输出目录", str(e))
            return

        try:
            dpi = int(self.var_dpi.get())
            quality = int(self.var_quality.get())
            crop_top = max(0, min(30, int(self.var_crop_top_pct.get())))
            crop_bot = max(0, min(30, int(self.var_crop_bottom_pct.get())))
            scan_band = max(1, min(40, int(self.var_scan_band.get())))
        except (tk.TclError, ValueError):
            messagebox.showerror("参数错误", "DPI / 质量 / 百分比 必须是整数。")
            return

        redact = RedactOptions(
            enabled=bool(self.var_redact_enabled.get()),
            smart=bool(self.var_redact_smart.get()),
            crop_top_pct=crop_top / 100.0,
            crop_bottom_pct=crop_bot / 100.0,
            scan_band=scan_band / 100.0,
        )
        fmt = self.var_fmt.get()

        self.cancel_event.clear()
        self.overall_done = 0
        self.var_progress.set(0.0)
        self.var_progress_text.set("0.0%")
        self.start_time = time.perf_counter()

        self._set_running(True)
        self._log("=" * 56)
        redact_desc = "否"
        if redact.enabled:
            parts = []
            if redact.smart:
                parts.append(f"智能(扫{int(redact.scan_band*100)}%)")
            if redact.crop_top_pct > 0:
                parts.append(f"裁页眉{int(redact.crop_top_pct*100)}%")
            if redact.crop_bottom_pct > 0:
                parts.append(f"裁页脚{int(redact.crop_bottom_pct*100)}%")
            redact_desc = ("是 [" + " + ".join(parts) + "]") if parts else "是 (无操作, 等同关闭)"
        self._log(f"开始转换  格式={fmt.upper()}  DPI={dpi}"
                  + (f"  JPG质量={quality}" if fmt == "jpeg" else "")
                  + f"  去除页码={redact_desc}")

        params = (list(self.scan_pdfs), input_dir, output_dir, dpi, fmt, quality, redact)
        self.worker = threading.Thread(target=self._run_conversion, args=params, daemon=True)
        self.worker.start()

    def _on_cancel(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            self.cancel_event.set()
            self.var_status.set("正在取消... (当前页结束后停止)")
            self._log("用户请求取消...")

    def _on_open_output(self) -> None:
        out = Path(self.var_output.get()).expanduser()
        if not out.exists():
            try:
                out.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("无法创建输出目录", str(e))
                return
        _open_in_explorer(out)

    def _on_close(self) -> None:
        try:
            self._save_settings()
        except Exception:
            pass
        if self.worker is not None and self.worker.is_alive():
            self.cancel_event.set()
            # 给 worker 一点时间退出, 避免线程卡死打包后的 exe
            self.worker.join(timeout=1.5)
        self.destroy()

    # ----- 后台 worker -----

    def _run_conversion(self, pdfs_with_pc, input_dir, output_dir, dpi, fmt, quality, redact):
        total_files = len(pdfs_with_pc)
        total_pages = sum(pc for _, pc in pdfs_with_pc) or 1
        all_results: List[FileResult] = []

        def page_cb(pdf_path: Path, page_no: int, total_in_pdf: int, redacted_n: int) -> None:
            self.overall_done += 1
            pct = self.overall_done / total_pages * 100.0
            self.q.put(("progress", (pct, self.overall_done, total_pages, pdf_path.name, page_no, total_in_pdf)))

        for idx, (pdf, _pc) in enumerate(pdfs_with_pc, start=1):
            if self.cancel_event.is_set():
                self.q.put(("log", "已取消, 终止后续文件。"))
                break
            try:
                rel = pdf.relative_to(input_dir)
            except ValueError:
                rel = pdf
            self.q.put(("log", f"[{idx}/{total_files}] 处理 {rel}"))
            try:
                result = convert_one_pdf(
                    pdf_path=pdf,
                    input_dir=input_dir,
                    output_dir=output_dir,
                    dpi=dpi,
                    fmt=fmt,
                    quality=quality,
                    redact=redact,
                    progress_cb=page_cb,
                    cancel_event=self.cancel_event,
                )
            except Exception as e:
                tb = traceback.format_exc()
                result = FileResult(pdf=pdf, errors=[f"转换异常: {e}", tb])
            all_results.append(result)
            tail = ""
            if redact.enabled and redact.smart:
                tail = f"  抹除页码 {result.redacted} 处"
            self.q.put(("log", f"    -> 成功 {result.ok} 页, 失败 {result.fail} 页{tail}"))
            for err in result.errors[:3]:
                self.q.put(("log", f"       * {err}"))

        elapsed = time.perf_counter() - self.start_time
        self.q.put(("done", (all_results, elapsed, output_dir, self.cancel_event.is_set())))

    # ----- 队列轮询 -----

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, data = self.q.get_nowait()
                self._handle_msg(kind, data)
        except queue.Empty:
            pass
        self.after(50, self._poll_queue)

    def _handle_msg(self, kind: str, data) -> None:
        if kind == "log":
            self._log(str(data))
        elif kind == "progress":
            pct, done, total, name, page_no, total_in_pdf = data
            self.var_progress.set(pct)
            self.var_progress_text.set(f"{pct:.1f}%")
            self.var_status.set(f"当前: {name}  第 {page_no}/{total_in_pdf} 页    总进度 {done}/{total} 页")
        elif kind == "done":
            results, elapsed, output_dir, cancelled = data
            self._on_done(results, elapsed, output_dir, cancelled)

    def _on_done(self, results: List[FileResult], elapsed: float, output_dir: Path, cancelled: bool) -> None:
        total_ok = sum(r.ok for r in results)
        total_fail = sum(r.fail for r in results)
        total_redacted = sum(r.redacted for r in results)
        self._log("=" * 56)
        if cancelled:
            self._log("已取消")
        else:
            self._log("全部完成")
        self._log(f"  PDF 文件 : {len(results)}")
        self._log(f"  成功页数 : {total_ok}")
        self._log(f"  失败页数 : {total_fail}")
        self._log(f"  抹除页码 : {total_redacted} 处")
        self._log(f"  总耗时   : {_format_seconds(elapsed)}")
        self._log(f"  输出位置 : {output_dir}")
        self.var_status.set(("已取消" if cancelled else "全部完成") + f"  共 {total_ok} 页, 失败 {total_fail} 页")
        self._set_running(False)
        if not cancelled and total_fail == 0 and total_ok > 0:
            self.var_progress.set(100.0)
            self.var_progress_text.set("100.0%")

    # ----- 工具 -----

    def _set_running(self, running: bool) -> None:
        if running:
            self.btn_start.configure(state="disabled")
            self.btn_scan.configure(state="disabled")
            self.btn_cancel.configure(state="normal")
        else:
            self.btn_start.configure(state="normal")
            self.btn_scan.configure(state="normal")
            self.btn_cancel.configure(state="disabled")

    def _log(self, msg: str) -> None:
        ts = time.strftime("%H:%M:%S")
        line = f"{ts}  {msg}\n"
        self.log.configure(state="normal")
        self.log.insert("end", line)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ----- 设置 -----

    def _apply_settings(self, s: Settings) -> None:
        base = app_dir()
        self.var_input.set(s.input_dir or str(base / "input"))
        self.var_output.set(s.output_dir or str(base / "output"))
        self.var_fmt.set(s.fmt if s.fmt in ("jpeg", "png") else "jpeg")
        idx_fmt = next((i for i, (_, v) in enumerate(FMT_OPTIONS) if v == self.var_fmt.get()), 0)
        try:
            self._cb_fmt.current(idx_fmt)
        except Exception:
            pass
        self.var_dpi.set(s.dpi if s.dpi in DPI_OPTIONS else 200)
        try:
            self._cb_dpi.set(str(self.var_dpi.get()))
        except Exception:
            pass
        self.var_quality.set(s.quality if s.quality in QUALITY_OPTIONS else 95)
        try:
            self._cb_q.set(str(self.var_quality.get()))
        except Exception:
            pass
        self.var_redact_enabled.set(bool(s.redact_enabled))
        self.var_redact_smart.set(bool(s.redact_smart))
        self.var_crop_top_pct.set(max(0, min(30, int(s.crop_top_pct))))
        self.var_crop_bottom_pct.set(max(0, min(30, int(s.crop_bottom_pct))))
        self.var_scan_band.set(max(5, min(30, int(s.scan_band))))
        self._update_redact_state()

    def _save_settings(self) -> None:
        s = Settings(
            input_dir=self.var_input.get().strip(),
            output_dir=self.var_output.get().strip(),
            fmt=self.var_fmt.get(),
            dpi=int(self.var_dpi.get() or 200),
            quality=int(self.var_quality.get() or 95),
            redact_enabled=bool(self.var_redact_enabled.get()),
            redact_smart=bool(self.var_redact_smart.get()),
            crop_top_pct=int(self.var_crop_top_pct.get() or 0),
            crop_bottom_pct=int(self.var_crop_bottom_pct.get() or 0),
            scan_band=int(self.var_scan_band.get() or 12),
        )
        s.save(self.settings_path)


def _format_seconds(s: float) -> str:
    if s < 60:
        return f"{s:.1f} 秒"
    m, sec = divmod(s, 60)
    return f"{int(m)} 分 {sec:.1f} 秒"


# ---------- 入口 ----------

def main() -> int:
    settings_path = app_dir() / SETTINGS_FILENAME
    settings = Settings.load(settings_path)
    if not settings.input_dir:
        (app_dir() / "input").mkdir(parents=True, exist_ok=True)
    if not settings.output_dir:
        (app_dir() / "output").mkdir(parents=True, exist_ok=True)

    app = App(settings, settings_path)
    app.mainloop()
    return 0


def _write_startup_log(message: str) -> None:
    """把启动期严重错误写到 exe 同目录, --windowed 下没有 stdout 的最后兜底。"""
    try:
        log_path = app_dir() / "PDFtoJPG_error.log"
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"\n--- {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
            f.write(message)
            f.write("\n")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        tb = traceback.format_exc()
        _write_startup_log(tb)
        try:
            _root = tk.Tk()
            _root.withdraw()
            messagebox.showerror("严重错误", f"程序异常终止:\n\n{tb}")
        except Exception:
            pass
        sys.exit(1)
