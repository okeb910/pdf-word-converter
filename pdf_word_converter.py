"""
PDF ↔ Word 互相转换工具

PDF → Word:
  - Microsoft Word 原生转换（推荐，质量通常较好，仍依赖本机 Word）
  - 整页渲染为图片嵌入 DOCX（视觉接近，文字一般不可编辑）
  - LibreOffice 无头转换（免费备选）

Word → PDF:
  - 优先 Microsoft Word COM，不可用时回退 LibreOffice
"""
import os
import io
import logging
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from app_environment import (
    APP_VERSION,
    build_install_command,
    find_winget,
    open_official_download,
    run_install_command,
)


# ═══════════════════════════════════════════════════════════
#  引擎可用性检测
# ═══════════════════════════════════════════════════════════

def _check_word_com_available():
    """检测 Microsoft Word 是否可通过 COM 使用"""
    try:
        import comtypes
        import comtypes.client
        comtypes.CoInitialize()
        try:
            word = comtypes.client.CreateObject("Word.Application", dynamic=True)
            word.Quit()
            return True
        finally:
            comtypes.CoUninitialize()
    except Exception:
        logging.exception("Microsoft Word COM detection failed")
        return False


def _find_libreoffice():
    """查找 LibreOffice 可执行文件路径"""
    import shutil
    # 1. 先检查 PATH 中的 libreoffice / soffice
    for name in ["libreoffice", "soffice"]:
        if shutil.which(name):
            return name
    # 2. 检查常见安装位置（Windows）
    import platform
    if platform.system() == "Windows":
        for base in [r"C:\Program Files\LibreOffice",
                     r"C:\Program Files (x86)\LibreOffice"]:
            soffice = os.path.join(base, "program", "soffice.exe")
            if os.path.exists(soffice):
                return soffice
    return None

_LO_PATH = None  # 缓存找到的 LibreOffice 路径

def _get_lo_path():
    global _LO_PATH
    if _LO_PATH is None:
        _LO_PATH = _find_libreoffice()
    return _LO_PATH

def _check_libreoffice_available():
    """用隔离的临时配置检测 LibreOffice 无头模式。"""
    import tempfile

    lo = _get_lo_path()
    if not lo:
        return False
    lo_dir = os.path.dirname(lo) or None
    try:
        with tempfile.TemporaryDirectory(prefix="pdf_word_converter_detect_") as profile_dir:
            profile_arg = f"-env:UserInstallation={Path(profile_dir).as_uri()}"
            result = subprocess.run(
                [lo, profile_arg, "--headless", "--terminate_after_init"],
                capture_output=True, text=True, timeout=20,
                cwd=lo_dir,
            )
        return result.returncode == 0
    except Exception:
        return False


def _load_pymupdf():
    """兼容 PyMuPDF 新旧模块名，并确认核心 API 存在。"""
    try:
        import pymupdf
        module = pymupdf
    except ImportError:
        import fitz
        module = fitz
    if not hasattr(module, "open"):
        raise ImportError("PyMuPDF 模块缺少 open()，请重新安装 PyMuPDF")
    return module


def _check_pymupdf_available():
    """检测 PyMuPDF 是否可用。"""
    try:
        _load_pymupdf()
        return True
    except ImportError:
        return False


# 缓存检测结果（加锁防竞态）
_ENGINE_LOCK = threading.Lock()
_WORD_AVAILABLE = None
_LIBREOFFICE_AVAILABLE = None
_PYMUPDF_AVAILABLE = None


def reset_engine_cache():
    """Clear engine detection state after an installation attempt."""
    global _LO_PATH, _WORD_AVAILABLE, _LIBREOFFICE_AVAILABLE, _PYMUPDF_AVAILABLE
    with _ENGINE_LOCK:
        _LO_PATH = None
        _WORD_AVAILABLE = None
        _LIBREOFFICE_AVAILABLE = None
        _PYMUPDF_AVAILABLE = None


def word_com_available():
    global _WORD_AVAILABLE
    with _ENGINE_LOCK:
        if _WORD_AVAILABLE is None:
            _WORD_AVAILABLE = _check_word_com_available()
    return _WORD_AVAILABLE


def libreoffice_available():
    global _LIBREOFFICE_AVAILABLE
    with _ENGINE_LOCK:
        if _LIBREOFFICE_AVAILABLE is None:
            _LIBREOFFICE_AVAILABLE = _check_libreoffice_available()
    return _LIBREOFFICE_AVAILABLE


def pymupdf_available():
    global _PYMUPDF_AVAILABLE
    with _ENGINE_LOCK:
        if _PYMUPDF_AVAILABLE is None:
            _PYMUPDF_AVAILABLE = _check_pymupdf_available()
    return _PYMUPDF_AVAILABLE


# ═══════════════════════════════════════════════════════════
#  PDF → Word 转换方法
# ═══════════════════════════════════════════════════════════

def pdf_to_word_via_word(pdf_path: str, docx_path: str, progress) -> None:
    """
    [推荐] Microsoft Word 原生转换
    Word 2013+ 可打开 PDF 并另存为可编辑 DOCX。
    质量通常优于纯第三方解析，但仍受 PDF 结构与本机 Word 版本影响。
    """
    import comtypes
    import comtypes.client
    import ctypes
    import time

    progress("正在启动 Microsoft Word...", 5)
    comtypes.CoInitialize()
    try:
        word = comtypes.client.CreateObject("Word.Application", dynamic=True)
    except Exception:
        comtypes.CoUninitialize()
        raise
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    # 后台线程：自动关闭 Word 可能弹出的 PDF 转换进度对话框
    stop_dismisser = False

    def dismiss_word_dialogs():
        """每隔 0.5 秒查找并关闭 Word 的转换进度弹窗"""
        user32 = ctypes.windll.user32
        WNDENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p,
        )

        triggers = [
            "正在转换", "Converting", "Opening PDF",
            "PDF 正在", "Word 正在",
        ]

        def enum_proc(hwnd, lParam):
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            if any(t in title for t in triggers):
                user32.PostMessageW(hwnd, 0x0100, 0x0D, 0)
                user32.PostMessageW(hwnd, 0x0101, 0x0D, 0)
            return True

        callback = WNDENUMPROC(enum_proc)
        while not stop_dismisser:
            user32.EnumWindows(callback, 0)
            time.sleep(0.5)

    dismisser_thread = threading.Thread(target=dismiss_word_dialogs, daemon=True)
    dismisser_thread.start()

    try:
        progress("正在打开 PDF 文件（Word 原生解析）...", 15)
        # Word 自动识别 PDF 格式，无需指定 Format 参数
        doc = word.Documents.Open(str(Path(pdf_path).absolute()))

        progress("正在转换为 Word 格式（保留完整排版）...", 50)
        # FileFormat=16 即 wdFormatDocumentDefault (.docx)
        doc.SaveAs2(str(Path(docx_path).absolute()), 16)
        doc.Close()
        progress("正在保存...", 90)
    finally:
        stop_dismisser = True
        dismisser_thread.join(timeout=1)
        word.Quit()
        comtypes.CoUninitialize()
    progress("完成", 100)


def pdf_to_word_via_images(pdf_path: str, docx_path: str, progress, dpi: int = 300) -> None:
    """
    [高保真观感] 每页 PDF 渲染为高清图片嵌入 DOCX
    视觉上通常非常接近原 PDF；文字一般不可编辑、不可检索。
    """
    fitz = _load_pymupdf()
    from docx import Document
    from docx.shared import Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    progress("正在读取 PDF...", 5)
    pdf_doc = fitz.open(pdf_path)
    total_pages = len(pdf_doc)

    word_doc = Document()

    for i in range(total_pages):
        progress(f"渲染第 {i + 1}/{total_pages} 页为图片...",
                  int(5 + (i / total_pages) * 80))

        page = pdf_doc[i]
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")

        page_w_inch = page.rect.width / 72
        page_h_inch = page.rect.height / 72

        # 获取或创建 section 和段落
        if i == 0:
            section = word_doc.sections[0]
            para = word_doc.add_paragraph()
        else:
            # add_section() 会同时创建新 section 和一个空段落，直接用该段落
            section = word_doc.add_section()
            para = word_doc.paragraphs[-1]

        section.page_width = Inches(page_w_inch)
        section.page_height = Inches(page_h_inch)
        section.top_margin = Cm(0)
        section.bottom_margin = Cm(0)
        section.left_margin = Cm(0)
        section.right_margin = Cm(0)

        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = para.paragraph_format
        pf.space_before = Cm(0)
        pf.space_after = Cm(0)
        pf.line_spacing = 1.0

        run = para.add_run()
        stream = io.BytesIO(img_bytes)
        run.add_picture(stream, width=Inches(page_w_inch))

    pdf_doc.close()

    progress("正在保存 Word 文档...", 90)
    word_doc.save(docx_path)
    progress("完成", 100)


def pdf_to_word_via_libreoffice(pdf_path: str, docx_path: str, progress) -> None:
    """通过 LibreOffice 无头模式转换。"""
    import shutil
    import tempfile

    lo = _get_lo_path()
    if not lo:
        raise RuntimeError(
            "LibreOffice 未安装或未找到。\n"
            "下载: https://www.libreoffice.org/download/"
        )
    progress("正在通过 LibreOffice 转换...", 10)
    lo_dir = os.path.dirname(lo) or None

    with tempfile.TemporaryDirectory(prefix="pdf_word_converter_") as temp_root:
        temp_root_path = Path(temp_root)
        temp_output_dir = temp_root_path / "output"
        profile_dir = temp_root_path / "profile"
        temp_output_dir.mkdir()
        profile_dir.mkdir()
        profile_arg = f"-env:UserInstallation={profile_dir.as_uri()}"
        result = subprocess.run(
            [
                lo, profile_arg, "--headless", "--convert-to", "docx",
                "--outdir", str(temp_output_dir), str(Path(pdf_path).absolute()),
            ],
            capture_output=True, text=True, timeout=300,
            cwd=lo_dir,
        )
        _raise_for_libreoffice_error(result, lo)
        generated = _find_libreoffice_output(
            temp_output_dir, Path(pdf_path).stem, ".docx"
        )
        shutil.move(str(generated), str(docx_path))
    progress("完成", 100)


def _raise_for_libreoffice_error(result, lo_path) -> None:
    stdout = getattr(result, "stdout", "") or ""
    stderr = getattr(result, "stderr", "") or ""
    combined_output = "\n".join(
        part.strip() for part in (stdout, stderr) if part.strip()
    )
    if result.returncode == 0 and "Error:" not in combined_output:
        return
    details = combined_output or f"LibreOffice 返回码 {result.returncode}，未提供详细信息"
    raise RuntimeError(
        f"LibreOffice 转换失败。\n"
        f"路径: {lo_path}\n"
        f"错误信息: {details}"
    )


def is_libreoffice_export_filter_error(error: Exception) -> bool:
    """Identify LibreOffice PDF import/export failures suitable for image fallback."""
    message = str(error).lower()
    indicators = (
        "export filter",
        "no export filter",
        "sfxbasemodel::impl_store",
        "未生成预期的输出文件",
        "source format: pdf",
    )
    return any(indicator in message for indicator in indicators)


def _find_libreoffice_output(temp_dir, source_stem, suffix) -> Path:
    expected = Path(temp_dir) / f"{source_stem}{suffix}"
    if expected.is_file():
        return expected
    candidates = list(Path(temp_dir).glob(f"*{suffix}"))
    if len(candidates) == 1:
        return candidates[0]
    raise RuntimeError("LibreOffice 未生成预期的输出文件")


# ═══════════════════════════════════════════════════════════
#  Word → PDF 转换方法
# ═══════════════════════════════════════════════════════════

def docx_to_pdf_via_word(docx_path: str, pdf_path: str, progress) -> None:
    """通过 Microsoft Word COM 导出 PDF"""
    import comtypes
    import comtypes.client

    progress("正在启动 Microsoft Word...", 5)
    comtypes.CoInitialize()
    try:
        word = comtypes.client.CreateObject("Word.Application", dynamic=True)
    except Exception:
        comtypes.CoUninitialize()
        raise
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    try:
        progress("正在打开 Word 文档...", 15)
        doc = word.Documents.Open(str(Path(docx_path).absolute()), False, True)

        progress("正在导出 PDF...", 50)
        # FileFormat=17 即 wdFormatPDF
        doc.SaveAs2(str(Path(pdf_path).absolute()), 17)
        doc.Close()
    finally:
        word.Quit()
        comtypes.CoUninitialize()
    progress("完成", 100)


def docx_to_pdf_via_libreoffice(docx_path: str, pdf_path: str, progress) -> None:
    """通过 LibreOffice 导出 PDF。"""
    import shutil
    import tempfile

    lo = _get_lo_path()
    if not lo:
        raise RuntimeError(
            "LibreOffice 未安装或未找到。\n"
            "下载: https://www.libreoffice.org/download/"
        )
    progress("正在通过 LibreOffice 导出...", 10)
    lo_dir = os.path.dirname(lo) or None

    with tempfile.TemporaryDirectory(prefix="pdf_word_converter_") as temp_root:
        temp_root_path = Path(temp_root)
        temp_output_dir = temp_root_path / "output"
        profile_dir = temp_root_path / "profile"
        temp_output_dir.mkdir()
        profile_dir.mkdir()
        profile_arg = f"-env:UserInstallation={profile_dir.as_uri()}"
        result = subprocess.run(
            [
                lo, profile_arg, "--headless", "--convert-to", "pdf",
                "--outdir", str(temp_output_dir), str(Path(docx_path).absolute()),
            ],
            capture_output=True, text=True, timeout=120,
            cwd=lo_dir,
        )
        _raise_for_libreoffice_error(result, lo)
        generated = _find_libreoffice_output(
            temp_output_dir, Path(docx_path).stem, ".pdf"
        )
        shutil.move(str(generated), str(pdf_path))
    progress("完成", 100)


# ═══════════════════════════════════════════════════════════
#  后处理：修复转换常见问题
# ═══════════════════════════════════════════════════════════

def fix_converted_docx(docx_path: str, progress=None) -> None:
    """对转换后的 DOCX 自动修复常见问题：
    1. 相邻单字母大小差异 → 小字母加下标
    2. 表格单元格字号统一
    3. 表格中去掉错误的 superscript
    4. 全角括号 （ ）→ 半角 ( ) 统一
    5. 短横 － 去上标
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from collections import Counter

    if progress:
        progress("正在自动修复格式...", 92)

    doc = Document(docx_path)

    # ── 1. 修复正文相邻单字母下标 ──
    for p in doc.paragraphs:
        runs = p.runs
        ri = 0
        while ri < len(runs) - 1:
            r1, r2 = runs[ri], runs[ri + 1]
            t1, t2 = r1.text.strip(), r2.text.strip()
            s1, s2 = r1.font.size, r2.font.size
            if (t1 and t2 and s1 and s2
                    and len(t1) == 1 and len(t2) == 1
                    and t1.isalpha() and t2.isalpha()):
                if 0.3 < s2 / s1 < 0.85:
                    # r2 is smaller → make it subscript
                    rpr = r2._element.find(qn('w:rPr'))
                    if rpr is None:
                        rpr = r2._element.makeelement(qn('w:rPr'), {})
                        r2._element.insert(0, rpr)
                    for old_va in rpr.findall(qn('w:vertAlign')):
                        rpr.remove(old_va)
                    va = rpr.makeelement(qn('w:vertAlign'), {qn('w:val'): 'subscript'})
                    rpr.append(va)
                    # Ensure r1 is not superscript
                    rpr1 = r1._element.find(qn('w:rPr'))
                    if rpr1 is not None:
                        for old_va in rpr1.findall(qn('w:vertAlign')):
                            rpr1.remove(old_va)
            ri += 1

    # ── 2 & 3 & 4 & 5. 修复表格 ──
    for table in doc.tables:
        # 收集列宽（用于统一）
        col_widths = []
        if table.rows:
            for ci, cell in enumerate(table.rows[0].cells):
                tcPr = cell._tc.find(qn('w:tcPr'))
                if tcPr is not None:
                    tcW = tcPr.find(qn('w:tcW'))
                    if tcW is not None:
                        col_widths.append((tcW.get(qn('w:w')), tcW.get(qn('w:type'))))
                    else:
                        col_widths.append(None)
                else:
                    col_widths.append(None)

        for row in table.rows:
            for ci, cell in enumerate(row.cells):
                # 收集字号并找到最优目标
                sizes = []
                all_runs = []
                for cp in cell.paragraphs:
                    for r in cp.runs:
                        if r.text.strip() and r.font.size:
                            sizes.append(r.font.size)
                            all_runs.append(r)

                if len(set(sizes)) > 1:
                    # 取最小号作为目标（大号通常是 pdf2docx 误放大）
                    target_size = min(sizes)
                    for r in all_runs:
                        r.font.size = target_size

                # 去 superscript、转换全角括号、修短横
                for cp in cell.paragraphs:
                    for r in cp.runs:
                        rpr = r._element.find(qn('w:rPr'))
                        # 去 superscript
                        if rpr is not None:
                            for va in list(rpr.findall(qn('w:vertAlign'))):
                                if va.get(qn('w:val')) == 'superscript':
                                    rpr.remove(va)
                        # 转全角括号
                        t_els = r._element.findall(qn('w:t'))
                        for t_el in t_els:
                            if t_el.text and ('（' in t_el.text or '）' in t_el.text):
                                t_el.text = t_el.text.replace('（', '(').replace('）', ')')
                        # 去短横上标
                        if r.text.strip() in ('－', '-', '–', '—'):
                            if rpr is not None:
                                for va in list(rpr.findall(qn('w:vertAlign'))):
                                    rpr.remove(va)

                # 统一列宽
                if ci < len(col_widths) and col_widths[ci] is not None:
                    tcPr = cell._tc.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = cell._tc.makeelement(qn('w:tcPr'), {})
                        cell._tc.insert(0, tcPr)
                    for old_tcw in tcPr.findall(qn('w:tcW')):
                        tcPr.remove(old_tcw)
                    w_val, w_type = col_widths[ci]
                    tcW = tcPr.makeelement(qn('w:tcW'), {qn('w:w'): w_val, qn('w:type'): w_type})
                    tcPr.append(tcW)

        # 统一表格内 C2 列对齐（去掉 JUSTIFY）
        for row in table.rows:
            if 2 < len(row.cells):
                for p in row.cells[2].paragraphs:
                    if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        p.alignment = None

    doc.save(docx_path)
    if progress:
        progress("格式修复完成", 95)


# ═══════════════════════════════════════════════════════════
#  GUI 界面
# ═══════════════════════════════════════════════════════════

from batch_logic import BatchResult, deduplicate_paths, run_conversion_batch


class ConverterApp:
    STATUS_TEXT = {
        "pending": "等待",
        "running": "转换中",
        "success": "成功",
        "failed": "失败",
        "cancelled": "已取消",
    }

    def __init__(self, root):
        self.root = root
        self.root.title(f"PDF ↔ Word 批量转换工具 v{APP_VERSION}")
        self.root.geometry("900x760")
        self.root.minsize(780, 680)
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

        self.input_paths = []
        self.batch_extension = None
        self.output_paths = []
        self._tree_paths = {}
        self._avail_methods = {}
        self._method_widgets = {}
        self._engine_detection_complete = False
        self._is_converting = False
        self._is_installing = False
        self._winget_path = None
        self._word_prompted = False
        self._cancel_event = threading.Event()
        self._setup_ui()
        self._detect_engines()

    def _setup_ui(self):
        title = tk.Label(
            self.root, text="PDF ↔ Word 批量转换工具",
            font=("Microsoft YaHei", 18, "bold"), bg="#f5f5f5", fg="#222",
        )
        title.pack(pady=(16, 4))

        tk.Label(
            self.root, text="本地串行转换 · 支持批量队列 · 文件不会上传",
            font=("Microsoft YaHei", 10), bg="#f5f5f5", fg="#666",
        ).pack(pady=(0, 12))

        file_frame = tk.LabelFrame(
            self.root, text=" 转换队列 ", font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=10, pady=8,
        )
        file_frame.pack(padx=24, fill="both")

        button_row = tk.Frame(file_frame, bg="#f5f5f5")
        button_row.pack(fill="x", pady=(0, 7))

        self.select_pdf_btn = tk.Button(
            button_row, text="添加 PDF", font=("Microsoft YaHei", 10),
            width=13, command=lambda: self.select_files(".pdf"),
            bg="#c0392b", fg="white", activebackground="#a93226", cursor="hand2",
        )
        self.select_pdf_btn.pack(side="left", padx=(0, 7))

        self.select_docx_btn = tk.Button(
            button_row, text="添加 Word", font=("Microsoft YaHei", 10),
            width=13, command=lambda: self.select_files(".docx"),
            bg="#2471a3", fg="white", activebackground="#1f618d", cursor="hand2",
        )
        self.select_docx_btn.pack(side="left", padx=(0, 7))

        self.remove_btn = tk.Button(
            button_row, text="移除选中", font=("Microsoft YaHei", 9),
            width=11, command=self.remove_selected,
        )
        self.remove_btn.pack(side="left", padx=(8, 5))

        self.clear_btn = tk.Button(
            button_row, text="清空", font=("Microsoft YaHei", 9),
            width=8, command=self.clear_queue,
        )
        self.clear_btn.pack(side="left")

        self.queue_summary_var = tk.StringVar(value="尚未添加文件")
        tk.Label(
            button_row, textvariable=self.queue_summary_var,
            font=("Microsoft YaHei", 9), bg="#f5f5f5", fg="#555",
        ).pack(side="right")

        tree_frame = tk.Frame(file_frame, bg="#f5f5f5")
        tree_frame.pack(fill="both", expand=True)
        columns = ("name", "folder", "status", "output")
        self.file_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=6,
            selectmode="extended",
        )
        self.file_tree.heading("name", text="文件名")
        self.file_tree.heading("folder", text="来源目录")
        self.file_tree.heading("status", text="状态")
        self.file_tree.heading("output", text="输出文件")
        self.file_tree.column("name", width=180, minwidth=120, stretch=False)
        self.file_tree.column("folder", width=250, minwidth=140, stretch=True)
        self.file_tree.column("status", width=75, minwidth=70, stretch=False, anchor="center")
        self.file_tree.column("output", width=260, minwidth=160, stretch=True)
        self.file_tree.tag_configure("success", foreground="#18794e")
        self.file_tree.tag_configure("failed", foreground="#b42318")
        self.file_tree.tag_configure("cancelled", foreground="#777")
        tree_scroll = ttk.Scrollbar(tree_frame, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y")
        self.file_tree.pack(side="left", fill="both", expand=True)
        self.file_tree.bind("<<TreeviewSelect>>", lambda _event: self._update_action_states())

        output_frame = tk.LabelFrame(
            self.root, text=" 输出位置 ", font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=10, pady=8,
        )
        output_frame.pack(padx=24, pady=(10, 0), fill="x")

        self.output_mode_var = tk.StringVar(value="source")
        self.source_output_radio = tk.Radiobutton(
            output_frame, text="各源文件所在目录", variable=self.output_mode_var,
            value="source", command=self._toggle_output_mode,
            font=("Microsoft YaHei", 9), bg="#f5f5f5",
        )
        self.source_output_radio.pack(side="left", padx=(0, 12))
        self.custom_output_radio = tk.Radiobutton(
            output_frame, text="统一输出目录", variable=self.output_mode_var,
            value="custom", command=self._toggle_output_mode,
            font=("Microsoft YaHei", 9), bg="#f5f5f5",
        )
        self.custom_output_radio.pack(side="left")

        self.output_dir_var = tk.StringVar()
        self.output_entry = tk.Entry(
            output_frame, textvariable=self.output_dir_var,
            font=("Consolas", 9), state="readonly", readonlybackground="white",
        )
        self.output_entry.pack(side="left", padx=8, fill="x", expand=True)
        self.browse_output_btn = tk.Button(
            output_frame, text="选择...", font=("Microsoft YaHei", 9),
            width=9, command=self.choose_output_dir,
        )
        self.browse_output_btn.pack(side="right")

        self.method_frame = tk.LabelFrame(
            self.root, text=" PDF → Word 转换方式 ", font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=10, pady=7,
        )
        self.method_frame.pack(padx=24, pady=(10, 0), fill="x")

        self.method_var = tk.StringVar(value="word_com")
        self._method_labels = {}
        methods_desc = [
            ("word_com", "Microsoft Word 原生转换（推荐，可编辑）"),
            ("images", "页面转高清图片嵌入（观感接近，不可编辑）"),
            ("libreoffice", "LibreOffice 引擎（兼容性有限）"),
        ]
        for value, description in methods_desc:
            row = tk.Frame(self.method_frame, bg="#f5f5f5")
            row.pack(anchor="w", pady=1, fill="x")
            radio = tk.Radiobutton(
                row, text=description, variable=self.method_var, value=value,
                font=("Microsoft YaHei", 9), bg="#f5f5f5", anchor="w",
                command=lambda selected=value: self._on_method_selected(selected),
            )
            radio.pack(side="left")
            status_label = tk.Label(
                row, text="检测中...", font=("Microsoft YaHei", 8),
                bg="#f5f5f5", fg="#777", width=10, anchor="e",
            )
            status_label.pack(side="right")
            self._method_widgets[value] = radio
            self._method_labels[value] = status_label

        action_frame = tk.Frame(self.root, bg="#f5f5f5")
        action_frame.pack(pady=(10, 7))
        self.convert_btn = tk.Button(
            action_frame, text="开始批量转换", font=("Microsoft YaHei", 10, "bold"),
            width=15, command=self.start_conversion,
            bg="#1e8449", fg="white", activebackground="#196f3d", cursor="hand2",
        )
        self.convert_btn.pack(side="left", padx=5)
        self.cancel_btn = tk.Button(
            action_frame, text="取消", font=("Microsoft YaHei", 10),
            width=10, command=self.cancel_conversion,
            bg="#a04000", fg="white", activebackground="#873600", cursor="hand2",
        )
        self.cancel_btn.pack(side="left", padx=5)
        self.open_btn = tk.Button(
            action_frame, text="打开输出目录", font=("Microsoft YaHei", 10),
            width=14, command=self.open_output_folder,
            bg="#555", fg="white", activebackground="#333", cursor="hand2",
        )
        self.open_btn.pack(side="left", padx=5)

        progress_frame = tk.Frame(self.root, bg="#f5f5f5")
        progress_frame.pack(padx=24, fill="x")
        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        self.progress_text_var = tk.StringVar(value="就绪")
        tk.Label(
            progress_frame, textvariable=self.progress_text_var,
            font=("Microsoft YaHei", 8), bg="#f5f5f5", fg="#666", anchor="w",
        ).pack(fill="x", pady=(2, 0))

        log_frame = tk.LabelFrame(
            self.root, text=" 转换日志 ", font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=7, pady=4,
        )
        log_frame.pack(padx=24, pady=(6, 12), fill="both", expand=True)
        self.log_text = tk.Text(
            log_frame, height=7, font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4", wrap="word", state="disabled",
        )
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        self._update_action_states()

    def _detect_engines(self, force=False):
        if force:
            reset_engine_cache()
        self._engine_detection_complete = False
        for label in self._method_labels.values():
            label.config(text="检测中...", fg="#777")
        self._update_action_states()

        def detect():
            have_word = word_com_available()
            have_libreoffice = libreoffice_available()
            have_pymupdf = pymupdf_available()
            winget_path = find_winget()
            self.root.after(
                0, self._update_engine_status,
                have_word, have_libreoffice, have_pymupdf, winget_path,
            )

        threading.Thread(target=detect, daemon=True).start()

    def _update_engine_status(self, have_word, have_libreoffice, have_pymupdf, winget_path):
        self._winget_path = winget_path
        self._avail_methods = {
            "word_com": have_word,
            "images": have_pymupdf,
            "libreoffice": have_libreoffice,
        }
        for value, label in self._method_labels.items():
            available = self._avail_methods[value]
            label.config(
                text="✓ 可用" if available else "✗ 不可用",
                fg="#18794e" if available else "#b42318",
            )

        if have_word:
            self.method_var.set("word_com")
        elif have_pymupdf:
            self.method_var.set("images")
        elif have_libreoffice:
            self.method_var.set("libreoffice")

        self._is_installing = False
        self._engine_detection_complete = True
        self._update_action_states()
        self.log(f"Microsoft Word: {'可用' if have_word else '不可用'}")
        self.log(f"PyMuPDF 图片模式: {'可用' if have_pymupdf else '不可用'}")
        self.log(f"LibreOffice: {'可用' if have_libreoffice else '不可用'}")
        self.log(f"winget: {'可用' if winget_path else '不可用'}")

        if not have_word and not self._word_prompted:
            self._word_prompted = True
            self.root.after(100, self._prompt_word_install)

    def _select_best_pdf_method(self):
        for method in ("word_com", "images", "libreoffice"):
            if self._avail_methods.get(method):
                self.method_var.set(method)
                return

    def _on_method_selected(self, method):
        if method == "libreoffice" and not self._avail_methods.get("libreoffice"):
            accepted = self._prompt_libreoffice_install()
            if not accepted:
                self._select_best_pdf_method()

    def _prompt_word_install(self):
        if self._avail_methods.get("word_com") or self._is_installing:
            return False
        confirmed = messagebox.askyesno(
            "未检测到 Microsoft Word",
            "是否使用 winget 安装 Microsoft Office？\n\n"
            "Office 需要有效许可证和 Microsoft 账号，安装过程可能要求管理员权限。"
            "程序不会自动提供许可证。",
        )
        if confirmed:
            self._start_install("office")
        else:
            self.log("用户已拒绝安装 Microsoft Office")
        return confirmed

    def _prompt_libreoffice_install(self):
        if self._avail_methods.get("libreoffice") or self._is_installing:
            return False
        confirmed = messagebox.askyesno(
            "需要 LibreOffice",
            "当前任务没有可用的 LibreOffice 引擎。是否使用 winget 安装官方 LibreOffice？",
        )
        if confirmed:
            self._start_install("libreoffice")
        else:
            self.log("用户已拒绝安装 LibreOffice")
        return confirmed

    def _start_install(self, product):
        display_name = "Microsoft Office" if product == "office" else "LibreOffice"
        if not self._winget_path:
            self.log(f"winget 不可用，正在打开 {display_name} 官方下载页")
            messagebox.showwarning(
                "winget 不可用",
                f"当前系统未找到 winget，将打开 {display_name} 官方下载页。",
            )
            open_official_download(product)
            return

        self._is_installing = True
        self._engine_detection_complete = False
        self._update_action_states()
        self.progress_text_var.set(f"正在安装 {display_name}...")
        self.log(f"开始安装 {display_name}")
        command = build_install_command(product, self._winget_path)

        def install():
            result = run_install_command(command)
            self.root.after(0, self._finish_install, product, display_name, result)

        threading.Thread(target=install, daemon=True).start()

    def _finish_install(self, product, display_name, result):
        if result.succeeded:
            self.log(f"{display_name} 安装命令已完成，正在重新检测引擎")
            messagebox.showinfo("安装完成", f"{display_name} 安装已完成，程序将立即重新检测。")
            self._detect_engines(force=True)
            return

        self._is_installing = False
        self._engine_detection_complete = True
        self._update_action_states()
        self.log(f"{display_name} 安装失败: {result.message}")
        messagebox.showwarning(
            "安装失败",
            f"{display_name} 安装未成功。将打开官方下载页。\n\n{result.message}",
        )
        open_official_download(product)

    def _ask_yes_no_from_worker(self, title, message):
        completed = threading.Event()
        answer = {"value": False}

        def ask():
            try:
                answer["value"] = messagebox.askyesno(title, message)
            finally:
                completed.set()

        self.root.after(0, ask)
        completed.wait()
        return answer["value"]

    def log(self, message: str):
        logging.info(message)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"  {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def select_files(self, extension: str):
        if extension == ".pdf":
            title = "添加 PDF 文件"
            filetypes = [("PDF 文件", "*.pdf")]
        else:
            title = "添加 Word 文件"
            filetypes = [("Word 文档", "*.docx")]

        selected = filedialog.askopenfilenames(title=title, filetypes=filetypes)
        if not selected:
            return

        selected_paths = [Path(path) for path in selected if Path(path).suffix.lower() == extension]
        if self.batch_extension and self.batch_extension != extension:
            confirmed = messagebox.askyesno(
                "切换转换方向",
                "当前队列包含另一种文件。是否清空现有队列并切换转换方向？",
            )
            if not confirmed:
                return
            self._clear_queue(log_change=False)

        previous_count = len(self.input_paths)
        self.input_paths = deduplicate_paths(self.input_paths + selected_paths)
        self.batch_extension = extension if self.input_paths else None
        self._refresh_queue()
        added_count = len(self.input_paths) - previous_count
        self.log(f"已添加 {added_count} 个文件；队列共 {len(self.input_paths)} 个")

    def _path_key(self, path: Path) -> str:
        return os.path.normcase(os.path.normpath(str(Path(path).absolute())))

    def _refresh_queue(self):
        self.file_tree.delete(*self.file_tree.get_children())
        self._tree_paths = {}
        for path in self.input_paths:
            item_id = self.file_tree.insert(
                "", "end",
                values=(path.name, str(path.parent), self.STATUS_TEXT["pending"], ""),
            )
            self._tree_paths[item_id] = path

        if self.batch_extension == ".pdf":
            direction = "PDF → Word"
        elif self.batch_extension == ".docx":
            direction = "Word → PDF"
        else:
            direction = ""
        summary = f"{direction} · {len(self.input_paths)} 个文件" if direction else "尚未添加文件"
        self.queue_summary_var.set(summary)
        self._update_action_states()

    def remove_selected(self):
        selected_ids = self.file_tree.selection()
        selected_keys = {
            self._path_key(self._tree_paths[item_id])
            for item_id in selected_ids if item_id in self._tree_paths
        }
        if not selected_keys:
            return
        self.input_paths = [
            path for path in self.input_paths if self._path_key(path) not in selected_keys
        ]
        if not self.input_paths:
            self.batch_extension = None
        self._refresh_queue()
        self.log(f"已移除 {len(selected_keys)} 个文件")

    def clear_queue(self):
        self._clear_queue(log_change=True)

    def _clear_queue(self, log_change):
        had_items = bool(self.input_paths)
        self.input_paths = []
        self.batch_extension = None
        self._refresh_queue()
        if had_items and log_change:
            self.log("已清空转换队列")

    def choose_output_dir(self):
        initial_dir = self.output_dir_var.get() or None
        directory = filedialog.askdirectory(title="选择统一输出目录", initialdir=initial_dir)
        if directory:
            self.output_dir_var.set(directory)
            self.output_mode_var.set("custom")
            self._toggle_output_mode()

    def _toggle_output_mode(self):
        self._update_action_states()

    def _validate_output_dir(self):
        if self.output_mode_var.get() == "source":
            return None
        raw_path = self.output_dir_var.get().strip()
        if not raw_path:
            raise ValueError("请选择统一输出目录")
        output_dir = Path(raw_path)
        if not output_dir.is_dir():
            raise ValueError("输出目录不存在")
        try:
            import tempfile
            with tempfile.NamedTemporaryFile(dir=str(output_dir), prefix=".pdf_word_converter_"):
                pass
        except OSError as exc:
            raise ValueError(f"输出目录不可写: {exc}") from exc
        return output_dir

    def start_conversion(self):
        if not self.input_paths:
            messagebox.showwarning("提示", "请先添加文件")
            return
        if not self._engine_detection_complete:
            messagebox.showinfo("提示", "转换引擎仍在检测，请稍候")
            return

        if self.batch_extension == ".pdf":
            method = self.method_var.get()
            if not self._avail_methods.get(method):
                if method == "libreoffice":
                    self._prompt_libreoffice_install()
                    return
                messagebox.showerror("转换方式不可用", "请选择一个当前可用的 PDF 转换方式")
                return
            target_suffix = ".docx"
        elif self.batch_extension == ".docx":
            method = None
            if not (self._avail_methods.get("word_com") or self._avail_methods.get("libreoffice")):
                self._prompt_libreoffice_install()
                return
            target_suffix = ".pdf"
        else:
            messagebox.showerror("队列错误", "队列中的文件类型不受支持")
            return

        try:
            output_dir = self._validate_output_dir()
        except ValueError as exc:
            messagebox.showerror("输出目录错误", str(exc))
            return

        self.output_paths = []
        self._cancel_event.clear()
        self._reset_queue_status()
        self._set_busy(True)
        self.progress["value"] = 0
        self.progress_text_var.set("准备转换...")
        direction = "PDF → Word" if self.batch_extension == ".pdf" else "Word → PDF"
        self.log(f"开始 {direction} 批量转换，共 {len(self.input_paths)} 个文件")

        paths = list(self.input_paths)
        extension = self.batch_extension
        threading.Thread(
            target=self._run_batch,
            args=(paths, extension, target_suffix, output_dir, method),
            daemon=True,
        ).start()

    def _reset_queue_status(self):
        for item_id in self.file_tree.get_children():
            values = list(self.file_tree.item(item_id, "values"))
            values[2] = self.STATUS_TEXT["pending"]
            values[3] = ""
            self.file_tree.item(item_id, values=values, tags=())

    def _run_batch(self, paths, extension, target_suffix, output_dir, method):
        converter = self._create_converter(extension, method)
        try:
            results = run_conversion_batch(
                paths,
                target_suffix,
                output_dir,
                converter,
                self._cancel_event,
                self._progress_update,
                self._status_update,
            )
            self.root.after(0, self._on_batch_complete, results)
        except Exception as exc:
            self.root.after(0, self._on_batch_error, str(exc))

    def _create_converter(self, extension, method):
        if extension == ".pdf":
            def convert_pdf(source, output, progress):
                if method == "word_com":
                    pdf_to_word_via_word(
                        source, output,
                        lambda message, pct: progress(message, min(88, int(pct * 0.88))),
                    )
                    fix_converted_docx(
                        output,
                        lambda message, pct: progress(message, min(99, max(89, pct))),
                    )
                elif method == "images":
                    pdf_to_word_via_images(source, output, progress)
                elif method == "libreoffice":
                    try:
                        pdf_to_word_via_libreoffice(
                            source, output,
                            lambda message, pct: progress(message, min(88, int(pct * 0.88))),
                        )
                    except Exception as exc:
                        can_retry = pymupdf_available() and is_libreoffice_export_filter_error(exc)
                        if not can_retry or not self._ask_yes_no_from_worker(
                            "LibreOffice 转换失败",
                            "LibreOffice 对 PDF → Word 的兼容性有限，当前文件无法导出。\n\n"
                            "是否改用内置图片模式重试？图片模式能保留页面观感，但文字不可编辑。",
                        ):
                            raise
                        output_path = Path(output)
                        if output_path.exists():
                            output_path.unlink()
                        progress("改用内置图片模式重试...", 5)
                        pdf_to_word_via_images(source, output, progress)
                        return
                    fix_converted_docx(
                        output,
                        lambda message, pct: progress(message, min(99, max(89, pct))),
                    )
                else:
                    raise RuntimeError(f"未知转换方式: {method}")

            return convert_pdf

        def convert_docx(source, output, progress):
            if word_com_available():
                docx_to_pdf_via_word(source, output, progress)
            elif libreoffice_available():
                docx_to_pdf_via_libreoffice(source, output, progress)
            else:
                raise RuntimeError("Word → PDF 需要 Microsoft Word 或 LibreOffice")

        return convert_docx

    def _progress_update(self, current, total, message, pct):
        self.root.after(0, self._set_progress, current, total, message, pct)

    def _set_progress(self, current, total, message, pct):
        self.progress["value"] = pct
        self.progress_text_var.set(f"[{current}/{total}] {message} · 总进度 {pct}%")
        self.log(f"[{current}/{total}] {message}")

    def _status_update(self, result: BatchResult):
        self.root.after(0, self._apply_status_update, result)

    def _apply_status_update(self, result: BatchResult):
        item_id = next(
            (
                current_id for current_id, path in self._tree_paths.items()
                if self._path_key(path) == self._path_key(result.source)
            ),
            None,
        )
        if item_id:
            values = list(self.file_tree.item(item_id, "values"))
            values[2] = self.STATUS_TEXT[result.status]
            values[3] = str(result.output) if result.output and result.status != "cancelled" else ""
            self.file_tree.item(item_id, values=values, tags=(result.status,))
            self.file_tree.see(item_id)
            item_number = self.file_tree.index(item_id) + 1
        else:
            item_number = 1

        total = max(1, len(self.input_paths))
        prefix = f"[{item_number}/{total}]"

        if result.status == "running":
            self.log(f"{prefix} 开始: {result.source.name}")
        elif result.status == "success":
            self.log(f"{prefix} 成功: {result.output}")
        elif result.status == "failed":
            self.log(f"{prefix} 失败: {result.source.name} - {result.error}")
        elif result.status == "cancelled":
            self.log(f"{prefix} 已取消: {result.source.name}")

    def cancel_conversion(self):
        if not self._is_converting or self._cancel_event.is_set():
            return
        self._cancel_event.set()
        self.cancel_btn.config(state="disabled")
        self.progress_text_var.set("正在等待当前文件处理结束后取消...")
        self.log("已请求取消；当前文件结束后将停止后续任务")

    def _on_batch_complete(self, results):
        self._set_busy(False)
        successful = [result for result in results if result.status == "success"]
        failed = [result for result in results if result.status == "failed"]
        cancelled = [result for result in results if result.status == "cancelled"]
        self.output_paths = [result.output for result in successful if result.output]

        summary = f"成功 {len(successful)}，失败 {len(failed)}，取消 {len(cancelled)}"
        self.queue_summary_var.set(f"批次完成 · {summary}")
        self.progress_text_var.set(summary)
        if not cancelled:
            self.progress["value"] = 100
        self.log(f"批次结束: {summary}")

        details = ""
        if failed:
            shown = failed[:3]
            details = "\n\n失败详情:\n" + "\n".join(
                f"- {result.source.name}: {result.error}" for result in shown
            )
            if len(failed) > len(shown):
                details += f"\n- 另有 {len(failed) - len(shown)} 个失败，请查看日志"

        message = summary + details
        if len(successful) == len(results):
            messagebox.showinfo("转换完成", message)
        elif not successful and failed and not cancelled:
            messagebox.showerror("转换失败", message)
        else:
            messagebox.showwarning("批次已结束", message)

    def _on_batch_error(self, error):
        self._set_busy(False)
        self.progress_text_var.set("批次异常终止")
        self.log(f"批次异常终止: {error}")
        messagebox.showerror("转换错误", error)

    def _set_busy(self, busy):
        self._is_converting = busy
        self._update_action_states()

    def _update_action_states(self):
        editable_state = "disabled" if self._is_converting or self._is_installing else "normal"
        self.select_pdf_btn.config(state=editable_state)
        self.select_docx_btn.config(state=editable_state)
        self.clear_btn.config(
            state="normal" if self.input_paths and not self._is_converting else "disabled"
        )
        self.remove_btn.config(
            state="normal" if self.file_tree.selection() and not self._is_converting else "disabled"
        )
        self.source_output_radio.config(state=editable_state)
        self.custom_output_radio.config(state=editable_state)
        browse_enabled = (
            not self._is_converting and not self._is_installing
            and self.output_mode_var.get() == "custom"
        )
        self.browse_output_btn.config(state="normal" if browse_enabled else "disabled")
        can_start = (
            self.input_paths and self._engine_detection_complete
            and not self._is_converting and not self._is_installing
        )
        self.convert_btn.config(state="normal" if can_start else "disabled")
        self.cancel_btn.config(state="normal" if self._is_converting else "disabled")

        for value, radio in self._method_widgets.items():
            enabled = (
                not self._is_converting and not self._is_installing
                and self.batch_extension != ".docx"
                and (self._avail_methods.get(value, False) or value == "libreoffice")
            )
            radio.config(state="normal" if enabled else "disabled")

    def open_output_folder(self):
        target = None
        if self.output_paths:
            target = self.output_paths[0].parent
        elif self.output_mode_var.get() == "custom" and self.output_dir_var.get():
            candidate = Path(self.output_dir_var.get())
            if candidate.is_dir():
                target = candidate
        elif self.input_paths:
            target = self.input_paths[0].parent

        if target and target.is_dir():
            os.startfile(str(target))
        else:
            messagebox.showinfo("提示", "当前没有可打开的输出目录")


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
