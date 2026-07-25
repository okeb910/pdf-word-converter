# PDF ↔ Word 转换工具

本地运行的 **PDF 与 Word（.docx）双向转换** 桌面小工具。  
优先调用本机 **Microsoft Word** 做转换，并提供 **整页图片嵌入**、**LibreOffice** 作为备选路径。

> 适合：本机批量处理、不希望文件上传到在线转换站的场景。  
> 不适合：追求「任意 PDF 一键变成可完美编辑 Word」——那在业界本身就很难保证。

![平台](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-yellow)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 功能概览

| 方向 | 能力 |
|------|------|
| **PDF → Word** | 三种方式可选（见下表） |
| **Word → PDF** | 自动选用本机 Word；不可用时回退 LibreOffice |
| **界面** | 文件选择、引擎可用性检测、进度条、转换日志 |
| **后处理** | 对「可编辑文字」类输出做常见格式修补（启发式，非万能） |

### PDF → Word 三种方式

| 方式 | 依赖 | 文字可编辑 | 版式/观感 | 说明 |
|------|------|------------|-----------|------|
| **Microsoft Word 原生** | 已安装的 Microsoft Word | 通常可以 | 依赖 Word 自身 PDF 导入能力 | **推荐**；质量通常最好，但仍可能因 PDF 结构而有偏差 |
| **页面转高清图片嵌入** | PyMuPDF | **否**（整页图） | 视觉高度一致 | 适合「看起来要对」、可接受不可选中文字 |
| **LibreOffice** | 本机 LibreOffice | 视引擎结果 | 一般 | 免费备选；部分版本需先手动启动一次完成初始化 |

Word → PDF 不单独选引擎：有 Word 用 Word，否则尝试 LibreOffice。

---

## 环境要求

- **操作系统**：Windows（当前实现依赖 Word COM / 常见 LO 路径；未做 macOS/Linux 适配）
- **Python**：3.8+
- **至少一种转换引擎**（按需求安装）：
  - [Microsoft Office / Word](https://www.microsoft.com/microsoft-365)（推荐，质量通常更好）
  - 和/或 [LibreOffice](https://www.libreoffice.org/download/)（免费）
  - 图片模式需要 Python 依赖 **PyMuPDF**（见 `requirements.txt`）

---

## 快速开始

### 1. 安装依赖

```bat
cd 本项目目录
pip install -r requirements.txt
```

### 2. 启动

双击：

```text
启动转换工具.bat
```

或：

```bat
python pdf_word_converter.py
```

### 3. 使用

1. 点 **「选择 PDF → Word」** 或 **「选择 Word → PDF」**
2. PDF→Word 时选择一种可用转换方式（界面会标注 ✓/✗）
3. 点 **「开始转换」**
4. 完成后可用 **「打开文件夹」** 查看输出

输出默认与源文件同目录：

- 若目标扩展名文件已存在，会生成 `原名_converted.docx` / `原名_converted.pdf`，避免直接覆盖。

---

## 项目结构

```text
.
├── pdf_word_converter.py   # 主程序（引擎检测 + 转换 + 后处理 + GUI）
├── requirements.txt        # Python 依赖
├── 启动转换工具.bat         # Windows 一键检查依赖并启动
├── LICENSE                 # MIT
├── README.md
└── .gitignore
```

当前是 **单文件应用**，便于阅读和二次修改；未拆成多包、未提供 CLI/HTTP 服务。

---

## 依赖说明

见 [`requirements.txt`](requirements.txt)：

| 包 | 用途 |
|----|------|
| `PyMuPDF` | PDF 渲染为图片（图片嵌入模式） |
| `python-docx` | 写入/修补 DOCX |
| `comtypes` | 调用 Microsoft Word COM |

Word / LibreOffice **不是** pip 包，需本机自行安装。

---

## 能力边界（请先读）

本工具 **包装并调度** 本机已有软件/库，而不是自研一套 PDF 排版引擎。因此：

1. **转换质量上限 ≈ 所选引擎上限**  
   复杂多栏、扫描件、特殊字体、表单域、严格印刷级 PDF，结果都可能不理想。

2. **「图片嵌入」≠ 可编辑文档**  
   观感接近原 PDF，但 Word 里基本是整页图，不能当正常文稿改字。

3. **后处理是启发式修补**  
   `fix_converted_docx` 会尝试处理：相邻单字母大小差异（作下标）、表格字号统一、去掉部分错误上标、全角括号转半角等。  
   这些规则来自实际踩坑，**不能保证**适合所有文档，也不应理解为「专业排版修复」。

4. **运行时会启动 Word / LibreOffice 进程**  
   转换中请勿强杀进程；异常时偶发残留 `WINWORD.EXE` 时，可在任务管理器中结束后再试。

5. **隐私**  
   转换在本地完成，本工具 **不会** 把文件上传到作者服务器。  
   （第三方软件 Word/LO 是否联网由其自身设置决定，与本仓库无关。）

---

## 和「万能在线转换」的定位差异

| | 本项目 | 常见在线转换站 |
|--|--------|----------------|
| 数据是否离开本机 | 默认本地 | 通常需上传 |
| 核心能力来源 | 本机 Word / LO / 渲染库 | 服务端引擎 |
| 开源可审计 | 是 | 多数否 |
| 跨平台开箱即用 | 目前偏 Windows | 浏览器即可 |
| 零依赖「完美可编辑」 | **不承诺** | 也很少真能承诺 |

如果你需要的是 **源格式可编辑**（例如本来就是 docx），应尽量保留源文件，而不是 PDF 来回倒。

---

## 开发说明

- 入口：`python pdf_word_converter.py`
- GUI：`tkinter`（Python 标准库）
- 转换在后台线程执行，通过 `root.after` 回写进度与日志
- 引擎检测结果带锁缓存，避免重复探测竞态

欢迎提 Issue / PR。若改引擎调用或后处理规则，建议附上 **输入样例特征**（可打码）和期望行为，便于回归。

---

## 路线图（非承诺）

以下为可能方向，**尚未实现**，仅作规划备忘：

- [ ] 批量转换、自定义输出目录
- [ ] 命令行接口（便于脚本/CI）
- [ ] 后处理规则可开关、可配置
- [ ] 更稳妥的 Word 进程生命周期管理
- [ ] 打包为免 Python 环境的发布包（如 PyInstaller）

---

## 许可证

本项目以 [MIT License](LICENSE) 发布。  
你可以使用、修改、分发；请保留版权与许可证声明。

第三方组件（Microsoft Word、LibreOffice、PyMuPDF 等）遵循其各自许可与使用条款。

---

## 免责声明

本软件按「现状」提供，作者不对转换结果的准确性、完整性或适用性作任何明示或暗示担保。  
用于合同、公文、试卷、财务等关键场景前，请 **人工核对** 输出文件。
