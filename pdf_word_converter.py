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
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path


# ═══════════════════════════════════════════════════════════
#  引擎可用性检测
# ═══════════════════════════════════════════════════════════

def _check_word_com_available():
    """检测 Microsoft Word 是否可通过 COM 使用"""
    try:
        import comtypes.client
        word = comtypes.client.CreateObject("Word.Application")
        word.Quit()
        return True
    except Exception:
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
    """检测 LibreOffice 是否可用（需要先手动运行一次 LO 完成初始化）"""
    lo = _get_lo_path()
    if not lo:
        return False
    lo_dir = os.path.dirname(lo)
    try:
        # 用 --headless 模式检测，确保能找到依赖库
        result = subprocess.run(
            [lo, "--headless", "--terminate_after_init"],
            capture_output=True, text=True, timeout=15,
            cwd=lo_dir,
        )
        # LO 26.x 返回码可能非零但库已初始化
        # 检查是否有 "platform independent libraries" 错误
        if "platform independent libraries" in (result.stderr or ""):
            return False
        return True
    except Exception:
        return False


def _check_pymupdf_available():
    """检测 PyMuPDF 是否可用"""
    try:
        import fitz
        return True
    except ImportError:
        return False


# 缓存检测结果（加锁防竞态）
_ENGINE_LOCK = threading.Lock()
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
    import comtypes.client
    import ctypes
    import time

    progress("正在启动 Microsoft Word...", 5)
    word = comtypes.client.CreateObject("Word.Application")
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
        doc = word.Documents.Open(
            str(Path(pdf_path).absolute()),
            ConfirmConversions=False,
        )

        progress("正在转换为 Word 格式（保留完整排版）...", 50)
        # FileFormat=16 即 wdFormatDocumentDefault (.docx)
        doc.SaveAs2(
            str(Path(docx_path).absolute()),
            FileFormat=16,
            CompatibilityMode=15,
        )
        doc.Close()
        progress("正在保存...", 90)
    finally:
        stop_dismisser = True
        dismisser_thread.join(timeout=1)
        word.Quit()
    progress("完成", 100)


def pdf_to_word_via_images(pdf_path: str, docx_path: str, progress, dpi: int = 300) -> None:
    """
    [高保真观感] 每页 PDF 渲染为高清图片嵌入 DOCX
    视觉上通常非常接近原 PDF；文字一般不可编辑、不可检索。
    """
    import fitz
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
    """通过 LibreOffice 无头模式转换"""
    lo = _get_lo_path()
    if not lo:
        raise RuntimeError(
            "LibreOffice 未安装或未找到。\n"
            "下载: https://www.libreoffice.org/download/"
        )
    progress("正在通过 LibreOffice 转换...", 10)
    out_dir = Path(docx_path).parent
    lo_dir = os.path.dirname(lo)
    lo_home = os.path.dirname(lo_dir)  # LO 安装根目录

    # 设置环境变量确保 LO 找到依赖库
    env = os.environ.copy()
    env["URE_BOOTSTRAP"] = f"file:///{lo_dir}/fundamental.ini"
    env["UNO_PATH"] = lo_dir
    # 将 LO 程序目录加入 PATH 以便加载 DLL
    env["PATH"] = f"{lo_dir};{lo_home}/share;{env.get('PATH', '')}"

    result = subprocess.run(
        [
            lo, "--headless", "--convert-to", "docx",
            "--outdir", str(out_dir), str(pdf_path),
        ],
        capture_output=True, text=True, timeout=300,
        cwd=lo_dir, env=env,
    )
    if result.returncode != 0:
        hint = ""
        if "platform independent libraries" in result.stderr:
            hint = (
                "\n\nLibreOffice 26.x 存在初始化问题。请尝试："
                "\n  1. 以管理员身份运行 LibreOffice 一次"
                "\n  2. 或安装 LibreOffice 24.x 稳定版"
            )
        raise RuntimeError(
            f"LibreOffice 转换失败。\n"
            f"路径: {lo}\n"
            f"错误信息: {result.stderr}{hint}"
        )
    # LO 生成的文件名可能与预期不同，查找实际输出
    lo_output = Path(out_dir) / (Path(pdf_path).stem + ".docx")
    if not lo_output.exists():
        # LO 可能用了不同的文件名
        for f in Path(out_dir).glob("*.docx"):
            if f.stem.startswith(Path(pdf_path).stem):
                lo_output = f
                break
    if lo_output != Path(docx_path) and lo_output.exists():
        import shutil
        if Path(docx_path).exists():
            Path(docx_path).unlink()
        shutil.move(str(lo_output), str(docx_path))
    progress("完成", 100)


# ═══════════════════════════════════════════════════════════
#  Word → PDF 转换方法
# ═══════════════════════════════════════════════════════════

def docx_to_pdf_via_word(docx_path: str, pdf_path: str, progress) -> None:
    """通过 Microsoft Word COM 导出 PDF"""
    import comtypes.client

    progress("正在启动 Microsoft Word...", 5)
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    try:
        progress("正在打开 Word 文档...", 15)
        doc = word.Documents.Open(str(Path(docx_path).absolute()), ReadOnly=True)

        progress("正在导出 PDF...", 50)
        # FileFormat=17 即 wdFormatPDF
        doc.SaveAs2(str(Path(pdf_path).absolute()), FileFormat=17)
        doc.Close()
    finally:
        word.Quit()
    progress("完成", 100)


def docx_to_pdf_via_libreoffice(docx_path: str, pdf_path: str, progress) -> None:
    """通过 LibreOffice 导出 PDF"""
    lo = _get_lo_path()
    if not lo:
        raise RuntimeError(
            "LibreOffice 未安装或未找到。\n"
            "下载: https://www.libreoffice.org/download/"
        )
    progress("正在通过 LibreOffice 导出...", 10)
    out_dir = Path(pdf_path).parent
    lo_dir = os.path.dirname(lo)
    lo_home = os.path.dirname(lo_dir)

    env = os.environ.copy()
    env["URE_BOOTSTRAP"] = f"file:///{lo_dir}/fundamental.ini"
    env["UNO_PATH"] = lo_dir
    env["PATH"] = f"{lo_dir};{lo_home}/share;{env.get('PATH', '')}"

    result = subprocess.run(
        [
            lo, "--headless", "--convert-to", "pdf",
            "--outdir", str(out_dir), str(docx_path),
        ],
        capture_output=True, text=True, timeout=120,
        cwd=lo_dir, env=env,
    )
    if result.returncode != 0:
        hint = ""
        if "platform independent libraries" in result.stderr:
            hint = (
                "\n\nLibreOffice 26.x 存在初始化问题。请尝试："
                "\n  1. 以管理员身份运行 LibreOffice 一次"
                "\n  2. 或安装 LibreOffice 24.x 稳定版"
            )
        raise RuntimeError(
            f"LibreOffice 转换失败。\n"
            f"路径: {lo}\n"
            f"错误信息: {result.stderr}{hint}"
        )
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

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF ↔ Word 转换工具（本地）")
        self.root.geometry("640x620")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f5f5")

        self.file_path = None
        self.output_path = None
        self._avail_methods = {}  # 存储每个方法的可用状态
        self._method_widgets = {}  # 存储 RadioButton 引用
        self._setup_ui()
        self._detect_engines()

    def _setup_ui(self):
        # ── 标题 ──
        title = tk.Label(
            self.root, text="PDF ↔ Word 转换工具",
            font=("Microsoft YaHei", 18, "bold"), bg="#f5f5f5", fg="#222",
        )
        title.pack(pady=(20, 5))

        tk.Label(
            self.root, text="本地转换 · 优先本机 Word · 可选图片保真 / LibreOffice",
            font=("Microsoft YaHei", 10), bg="#f5f5f5", fg="#888",
        ).pack(pady=(0, 18))

        # ── 文件选择 ──
        file_frame = tk.LabelFrame(
            self.root, text=" 选择文件 ", font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=12, pady=10,
        )
        file_frame.pack(padx=30, fill="x")

        btn_row = tk.Frame(file_frame, bg="#f5f5f5")
        btn_row.pack(fill="x")

        self.select_pdf_btn = tk.Button(
            btn_row, text="选择 PDF → Word", font=("Microsoft YaHei", 10),
            width=17, command=lambda: self.select_file("pdf"),
            bg="#e74c3c", fg="white", activebackground="#c0392b", cursor="hand2",
        )
        self.select_pdf_btn.pack(side="left", padx=(0, 8))

        self.select_docx_btn = tk.Button(
            btn_row, text="选择 Word → PDF", font=("Microsoft YaHei", 10),
            width=17, command=lambda: self.select_file("docx"),
            bg="#2980b9", fg="white", activebackground="#1a5276", cursor="hand2",
        )
        self.select_docx_btn.pack(side="left")

        self.path_var = tk.StringVar(value="尚未选择文件")
        path_entry = tk.Entry(
            file_frame, textvariable=self.path_var, font=("Consolas", 9),
            state="readonly", readonlybackground="white",
        )
        path_entry.pack(fill="x", pady=(10, 0))

        # ── PDF → Word 转换方式 ──
        self.method_frame = tk.LabelFrame(
            self.root, text=" PDF → Word 转换方式 ",
            font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=12, pady=8,
        )
        self.method_frame.pack(padx=30, pady=(12, 8), fill="x")

        self.method_var = tk.StringVar(value="word_com")

        # 各方法的描述
        methods_desc = [
            ("word_com",  "Microsoft Word 原生转换（推荐，可编辑，质量视 PDF 而定）"),
            ("images",   "页面转高清图片嵌入（观感接近，文字通常不可编辑）"),
            ("libreoffice", "LibreOffice 引擎（免费备选）"),
        ]
        self._method_labels = {}  # 存储 Label 引用用于更新状态

        for val, desc in methods_desc:
            row = tk.Frame(self.method_frame, bg="#f5f5f5")
            row.pack(anchor="w", pady=1, fill="x")

            rb = tk.Radiobutton(
                row, text=desc, variable=self.method_var, value=val,
                font=("Microsoft YaHei", 9), bg="#f5f5f5", anchor="w",
            )
            rb.pack(side="left")

            status_lbl = tk.Label(
                row, text="检测中...", font=("Microsoft YaHei", 8),
                bg="#f5f5f5", fg="#999", width=12, anchor="e",
            )
            status_lbl.pack(side="right")

            self._method_widgets[val] = rb
            self._method_labels[val] = status_lbl

        # ── 按钮区 ──
        btn_frame = tk.Frame(self.root, bg="#f5f5f5")
        btn_frame.pack(pady=(8, 10))

        self.convert_btn = tk.Button(
            btn_frame, text="开始转换", font=("Microsoft YaHei", 11, "bold"),
            width=15, command=self.start_conversion,
            bg="#27ae60", fg="white", activebackground="#1e8449", cursor="hand2",
        )
        self.convert_btn.pack(side="left", padx=6)

        self.open_btn = tk.Button(
            btn_frame, text="打开文件夹", font=("Microsoft YaHei", 11),
            width=15, command=self.open_output_folder,
            bg="#555", fg="white", activebackground="#333", cursor="hand2",
        )
        self.open_btn.pack(side="left", padx=6)

        # ── 进度条 ──
        self.progress = ttk.Progressbar(
            self.root, mode="determinate", length=480, maximum=100,
        )
        self.progress.pack(pady=(8, 4))

        # ── 日志 ──
        log_frame = tk.LabelFrame(
            self.root, text=" 转换日志 ", font=("Microsoft YaHei", 10),
            bg="#f5f5f5", fg="#333", padx=8, pady=4,
        )
        log_frame.pack(padx=30, pady=(4, 10), fill="both", expand=True)

        text_scroll_frame = tk.Frame(log_frame, bg="#f5f5f5")
        text_scroll_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            text_scroll_frame, height=8, font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4", wrap="word", state="disabled",
        )
        scrollbar = ttk.Scrollbar(text_scroll_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

    def _detect_engines(self):
        """在后台检测可用引擎并更新 UI"""
        def _detect():
            have_word = word_com_available()
            have_libre = libreoffice_available()
            have_pymupdf = pymupdf_available()
            self.root.after(0, self._update_engine_status,
                            have_word, have_libre, have_pymupdf)

        threading.Thread(target=_detect, daemon=True).start()

    def _update_engine_status(self, have_word, have_libre, have_pymupdf):
        self._avail_methods = {
            "word_com": have_word,
            "images": have_pymupdf,
            "libreoffice": have_libre,
        }

        for val, label in self._method_labels.items():
            if self._avail_methods.get(val):
                label.config(text="✓ 可用", fg="#27ae60")
            else:
                label.config(text="✗ 不可用", fg="#e74c3c")

        # 自动选择最佳可用方法
        if have_word:
            self.method_var.set("word_com")
        elif have_pymupdf:
            self.method_var.set("images")
        else:
            self.method_var.set("libreoffice")

        for val, rb in self._method_widgets.items():
            if not self._avail_methods.get(val):
                rb.config(state="disabled")

        if have_word:
            self.log("✓ 检测到 Microsoft Word — 可使用原生高质量转换")
        else:
            self.log("✗ 未检测到 Microsoft Word")
        if have_pymupdf:
            self.log("✓ 检测到 PyMuPDF — 图片嵌入模式可用")
        else:
            self.log("✗ 未检测到 PyMuPDF，图片模式不可用")
        if have_libre:
            self.log("✓ 检测到 LibreOffice")
        else:
            self.log("✗ 未检测到 LibreOffice（可选）")

    def log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"  {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def select_file(self, ext_hint: str):
        filetypes = [
            ("支持的文件", "*.pdf;*.docx"),
            ("PDF 文件", "*.pdf"),
            ("Word 文档", "*.docx"),
        ]
        file_path = filedialog.askopenfilename(title="选择文件", filetypes=filetypes)
        if not file_path:
            return

        ext = Path(file_path).suffix.lower()
        if ext not in (".pdf", ".docx"):
            messagebox.showwarning("提示", f"不支持的文件类型: {ext}")
            return

        self.file_path = file_path
        self.path_var.set(file_path)
        self.log(f"已选择: {os.path.basename(file_path)}")

        # 根据文件类型显示相关提示
        if ext == ".pdf":
            self.log(f"转换方向: PDF → Word")
        else:
            self.log(f"转换方向: Word → PDF")

    def start_conversion(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先选择文件")
            return
        if not os.path.exists(self.file_path):
            messagebox.showerror("错误", "文件不存在")
            return

        ext = Path(self.file_path).suffix.lower()

        self.convert_btn["state"] = "disabled"
        self.select_pdf_btn["state"] = "disabled"
        self.select_docx_btn["state"] = "disabled"
        self.progress["value"] = 0

        if ext == ".pdf":
            method = self.method_var.get()
            self.log(f"开始 PDF → Word 转换（方式: {method}）")
            threading.Thread(
                target=self._run_pdf_to_word, args=(method,), daemon=True
            ).start()
        elif ext == ".docx":
            self.log("开始 Word → PDF 转换...")
            threading.Thread(target=self._run_docx_to_pdf, daemon=True).start()

    def _run_pdf_to_word(self, method):
        try:
            input_path = Path(self.file_path)
            output_path = input_path.with_suffix(".docx")
            if output_path.exists():
                output_path = input_path.with_name(
                    f"{input_path.stem}_converted.docx"
                )

            pdf_path = str(input_path)
            docx_path = str(output_path)

            if method == "word_com":
                pdf_to_word_via_word(pdf_path, docx_path, self._progress_update)
                # Word COM 输出也做后处理
                fix_converted_docx(docx_path, self._progress_update)
            elif method == "images":
                pdf_to_word_via_images(pdf_path, docx_path, self._progress_update)
                # 图片模式不需要后处理（整页图片，无文字格式问题）
            elif method == "libreoffice":
                pdf_to_word_via_libreoffice(pdf_path, docx_path, self._progress_update)
                # LO 输出也做后处理
                fix_converted_docx(docx_path, self._progress_update)
            else:
                raise RuntimeError(f"未知转换方式: {method}")

            self.output_path = str(output_path)
            self.root.after(0, self._on_success,
                            "PDF → Word 转换完成！", str(output_path))
        except Exception as e:
            import traceback
            detail = traceback.format_exc()
            self.root.after(0, self._on_error, f"转换失败: {e}\n\n{detail}")

    def _run_docx_to_pdf(self):
        try:
            input_path = Path(self.file_path)
            output_path = input_path.with_suffix(".pdf")
            if output_path.exists():
                output_path = input_path.with_name(
                    f"{input_path.stem}_converted.pdf"
                )

            docx_path = str(input_path)
            pdf_path = str(output_path)

            if word_com_available():
                docx_to_pdf_via_word(docx_path, pdf_path, self._progress_update)
            elif libreoffice_available():
                docx_to_pdf_via_libreoffice(docx_path, pdf_path, self._progress_update)
            else:
                raise RuntimeError(
                    "Word → PDF 需要 Microsoft Word 或 LibreOffice。\n"
                    "请安装其中之一：\n"
                    "  • Microsoft Office (推荐)\n"
                    "  • LibreOffice (免费): https://www.libreoffice.org/"
                )

            self.output_path = str(output_path)
            self.root.after(0, self._on_success,
                            "Word → PDF 转换完成！", str(output_path))
        except Exception as e:
            import traceback
            detail = traceback.format_exc()
            self.root.after(0, self._on_error, f"转换失败: {e}\n\n{detail}")

    def _progress_update(self, msg: str, pct: int):
        self.root.after(0, self._set_progress, msg, pct)

    def _set_progress(self, msg, pct):
        self.log(msg)
        self.progress["value"] = pct

    def _on_success(self, title, path):
        self.progress["value"] = 100
        self.convert_btn["state"] = "normal"
        self.select_pdf_btn["state"] = "normal"
        self.select_docx_btn["state"] = "normal"
        self.log(f"✓ {title}")
        self.log(f"  输出文件: {path}")
        messagebox.showinfo("转换完成", f"{title}\n\n输出文件:\n{path}")

    def _on_error(self, msg):
        self.progress["value"] = 0
        self.convert_btn["state"] = "normal"
        self.select_pdf_btn["state"] = "normal"
        self.select_docx_btn["state"] = "normal"
        self.log(f"✗ {msg}")
        messagebox.showerror("转换错误", msg)

    def open_output_folder(self):
        target = None
        if self.output_path and os.path.exists(self.output_path):
            target = os.path.dirname(self.output_path)
        elif self.file_path and os.path.exists(self.file_path):
            target = os.path.dirname(self.file_path)
        if target:
            os.startfile(target)
        else:
            messagebox.showinfo("提示", "尚未进行过转换操作")


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
