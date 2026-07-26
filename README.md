# PDF ↔ Word 批量转换工具

## v0.3.0 免 Python 版本

v0.3.0 同时提供 64 位 Windows 10/11 的便携单文件版和安装程序版。两个版本都内置 Python 3.12、Tkinter、PyMuPDF、python-docx 和 comtypes，目标电脑不需要安装 Python，启动时也不会执行 pip。

- `PDF-Word批量转换工具-v0.3.0-便携版-x64.exe`：直接双击运行。
- `PDF-Word批量转换工具-v0.3.0-安装版-x64.exe`：按当前用户安装到 `%LOCALAPPDATA%\Programs\PDFWordConverter`，带开始菜单快捷方式和卸载项。
- 每次启动都会在后台检测内置组件、Microsoft Word、LibreOffice 和 winget，检测过程不阻塞主窗口。
- 未检测到 Word 时，每次进程最多询问一次是否通过 winget 安装 Microsoft Office；Office 仍需要有效许可证和 Microsoft 账号，并可能要求管理员权限。
- LibreOffice 不会在启动时主动提示，只在选择 LibreOffice 或 Word → PDF 没有可用引擎时询问安装。
- 安装命令只会在用户确认后以参数列表执行；winget 不可用或安装失败时打开官方页面。
- 启动和环境检测日志位于 `%LOCALAPPDATA%\PDFWordConverter\logs`。

LibreOffice 的 PDF → Word 支持标记为“兼容性有限”。遇到无导出过滤器时，程序会询问是否改用内置图片模式重试。

### 本地构建

构建机需要 Python 3.12 x64 和 Inno Setup 6；这些工具不会安装到目标电脑。固定版本依赖见 `requirements-build.txt`：

```powershell
.\build_release.ps1 -PythonExe "C:\path\to\python.exe" -ReleaseDir "C:\path\to\release"
```

PyInstaller 配置、Windows manifest、版本资源、简体中文安装语言文件和 Inno Setup 脚本位于 `packaging/`。

---

本地运行的 **PDF 与 Word（.docx）双向批量转换** 桌面工具。
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
| **批量队列** | 一次添加多个同类型文件，逐项显示状态与输出路径 |
| **输出管理** | 默认输出到源目录，也可统一指定目录；自动编号且不覆盖文件 |
| **执行控制** | 总进度、逐项日志、失败后继续、安全取消和批次结果汇总 |
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

### 3. 批量转换

1. 点 **「添加 PDF」** 或 **「添加 Word」**，可一次多选文件。
2. 继续点击同类按钮可追加文件；重复路径会自动去重。
3. PDF → Word 时选择一种可用转换方式。
4. 保持输出到各源文件目录，或选择一个统一输出目录。
5. 点 **「开始批量转换」**，在队列、总进度和日志中查看结果。
6. 需要停止时点 **「取消」**；程序会先结束当前文件，再取消剩余项目。

同一批次只处理一种方向。队列已有另一种文件时，程序会先询问是否清空并切换方向。

### 输出与重名规则

默认输出到每个源文件所在目录，也可以把整批结果输出到同一个自定义目录。

程序绝不主动覆盖已有结果。若文件已存在，会依次生成：

```text
原名.docx
原名_converted.docx
原名_converted_2.docx
原名_converted_3.docx
```

Word → PDF 使用相同编号规则，只是扩展名改为 `.pdf`。

---

## 项目结构

```text
.
├── pdf_word_converter.py   # 转换引擎、后处理与 GUI
├── batch_logic.py          # 输出编号与可测试的串行批次调度
├── app_environment.py      # 冻结环境检测与按需安装命令
├── launcher.py             # 打包版启动检查与错误日志
├── build_release.ps1       # 本地 v0.3.0 构建脚本
├── requirements-build.txt  # Python 3.12 固定构建依赖
├── packaging/              # PyInstaller、manifest、版本资源、Inno Setup
├── tests/                  # 批次逻辑自动化测试
├── requirements.txt        # Python 依赖
├── 启动转换工具.bat         # Windows 一键检查依赖并启动
├── LICENSE                 # MIT
└── README.md
```

当前仍是轻量桌面应用；未提供 CLI/HTTP 服务。

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
   `fix_converted_docx` 会尝试处理相邻单字母大小差异、表格字号、错误上标、全角括号等问题。这些规则不能保证适合所有文档。

4. **取消不会强制终止当前转换**
   为避免损坏文件或留下异常 Office 进程，取消请求会在当前文件完成或失败后生效。

5. **运行时会启动 Word / LibreOffice 进程**
   转换中请勿强杀进程；异常时偶发残留 `WINWORD.EXE` 时，可在任务管理器中结束后再试。

6. **隐私**
   转换在本地完成，本工具 **不会** 把文件上传到作者服务器。第三方软件是否联网由其自身设置决定。

---

## 开发说明

- 入口：`python pdf_word_converter.py`
- GUI：`tkinter`（Python 标准库）
- 批次逻辑：`batch_logic.py`，与界面和具体转换引擎解耦
- 转换在后台线程串行执行，通过 `root.after` 回写队列、进度与日志
- 引擎检测结果带锁缓存，避免重复探测竞态
- 自动化测试：`python -m unittest discover -s tests -v`

若修改转换引擎或后处理规则，建议附上输入样例特征（可打码）和期望行为，便于回归。

---

## 路线图（非承诺）

- [x] 批量转换、自定义输出目录
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
