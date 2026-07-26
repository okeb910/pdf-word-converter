# PDF ↔ Word 批量转换工具

Windows 本地 **PDF 与 Word（.docx）双向批量转换** 桌面工具。  
优先调用本机 **Microsoft Word**，并提供 **整页图片嵌入**、**LibreOffice** 作为备选。

> 适合：本机批量处理、不希望文件上传到在线转换站  
> 不适合：追求「任意 PDF 一键变成可完美编辑 Word」——业界本身也很难保证

![平台](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-yellow)
![License](https://img.shields.io/badge/license-MIT-green)
![Release](https://img.shields.io/github/v/release/okeb910/pdf-word-converter)

---

## 普通用户：下载安装包（推荐）

**无需安装 Python。** 请从 Releases 下载：

**[→ 下载 v0.3.0 安装包](https://github.com/okeb910/pdf-word-converter/releases/tag/v0.3.0)**

| 文件 | 说明 |
|------|------|
| `PDFWordConverter-v0.3.0-setup-x64.exe` | **安装版**（推荐）：安装到当前用户目录，带开始菜单与卸载项 |
| `PDFWordConverter-v0.3.0-portable-x64.exe` | **便携版**：双击即用，无需安装 |
| `SHA256SUMS.txt` | 可选：校验文件完整性 |

- 系统：**64 位 Windows 10 / 11**
- 推荐本机安装 [Microsoft Word](https://www.microsoft.com/microsoft-365)（转换质量通常最好）
- 程序可能提示「未知发布者」（未数字签名）：点「更多信息」→「仍要运行」
- 启动与环境检测日志：`%LOCALAPPDATA%\PDFWordConverter\logs`

### 三分钟上手

1. 运行安装版或便携版（二选一）
2. 点 **「添加 PDF」** 或 **「添加 Word」**（可多选）
3. PDF → Word 时选择一种转换方式（见下表）
4. 选择输出目录（默认与源文件同目录）
5. 点 **「开始批量转换」**

### PDF → Word 三种方式

| 方式 | 依赖 | 文字可编辑 | 说明 |
|------|------|------------|------|
| **Microsoft Word 原生** | 本机 Word | 通常可以 | **推荐**，质量通常最好 |
| **页面转高清图片嵌入** | 程序自带 | **否**（整页图） | 版式接近原 PDF，不能当正文改字 |
| **LibreOffice** | 本机 LibreOffice | 视结果 | 免费备选；兼容性有限 |

Word → PDF：有 Word 用 Word，否则尝试 LibreOffice。

### 输出命名规则

不会覆盖已有文件，依次生成：

```text
原名.docx
原名_converted.docx
原名_converted_2.docx
...
```

---

## 功能概览

| 能力 | 说明 |
|------|------|
| **PDF → Word** | 三种方式可选 |
| **Word → PDF** | 自动选用 Word，不可用时回退 LibreOffice |
| **批量队列** | 一次添加多个同类型文件，逐项显示状态与输出路径 |
| **输出管理** | 默认源目录，或统一输出目录；自动编号不覆盖 |
| **执行控制** | 总进度、日志、失败后继续、安全取消 |
| **后处理** | 对可编辑文字类输出做常见格式修补（启发式，非万能） |

未检测到 Word 时，程序可能询问是否通过 winget 安装 Office（需你确认；Office 仍需许可证）。  
LibreOffice 仅在你选择该引擎或 Word → PDF 无可用引擎时询问安装。

---

## 开发者：从源码运行

### 环境要求

- Windows
- Python 3.8+（建议 3.12）
- 至少一种转换引擎：Microsoft Word 和/或 LibreOffice；图片模式另需 PyMuPDF

### 安装与启动

```bat
git clone https://github.com/okeb910/pdf-word-converter.git
cd pdf-word-converter
pip install -r requirements.txt
```

双击 `启动转换工具.bat`，或：

```bat
python pdf_word_converter.py
```

### 项目结构

```text
.
├── pdf_word_converter.py   # 转换引擎、后处理与 GUI
├── batch_logic.py          # 输出编号与串行批次调度
├── app_environment.py      # 冻结环境检测与按需安装命令
├── launcher.py             # 打包版启动检查与错误日志
├── build_release.ps1       # 本地构建脚本
├── requirements-build.txt  # 构建依赖（固定版本）
├── packaging/              # PyInstaller、manifest、Inno Setup
├── tests/                  # 自动化测试
├── requirements.txt        # 运行依赖
├── 启动转换工具.bat
├── LICENSE
└── README.md
```

### 依赖说明

| 包 | 用途 |
|----|------|
| `PyMuPDF` | PDF 渲染为图片（图片嵌入模式） |
| `python-docx` | 写入/修补 DOCX |
| `comtypes` | 调用 Microsoft Word COM |

Word / LibreOffice **不是** pip 包，需本机自行安装。

### 本地构建安装包

构建机需要 **Python 3.12 x64** 和 **Inno Setup 6**（不会装到用户电脑）：

```powershell
.\build_release.ps1 -PythonExe "C:\path\to\python.exe" -ReleaseDir "C:\path\to\release"
```

配置见 `packaging/`。

### 测试

```bat
python -m unittest discover -s tests -v
```

---

## 能力边界（请先读）

本工具 **包装并调度** 本机已有软件/库，而不是自研 PDF 排版引擎：

1. **质量上限 ≈ 所选引擎上限** — 复杂多栏、扫描件、特殊字体等可能不理想  
2. **图片嵌入 ≠ 可编辑文档** — 观感接近原 PDF，但基本不能改字  
3. **后处理是启发式修补** — 不保证适合所有文档  
4. **取消不会强杀当前转换** — 等当前文件完成或失败后再停，避免损坏文件  
5. **运行时会启动 Word / LibreOffice** — 请勿强杀进程；异常残留可在任务管理器结束 `WINWORD.EXE`  
6. **隐私** — 转换在本地完成，本工具不会把文件上传到作者服务器  

---

## 许可证

本项目以 [MIT License](LICENSE) 发布。  
你可以使用、修改、分发；请保留版权与许可证声明。

第三方组件（Microsoft Word、LibreOffice、PyMuPDF 等）遵循其各自许可与使用条款。

---

## 免责声明

本软件按「现状」提供，作者不对转换结果的准确性、完整性或适用性作任何明示或暗示担保。  
用于合同、公文、试卷、财务等关键场景前，请 **人工核对** 输出文件。
