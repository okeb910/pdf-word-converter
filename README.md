# PDF ↔ Word/PPT 批量转换工具

当前版本为 **v0.5.0**，由同一套源码支持 Windows 与 macOS；Windows 面向普通用户提供免 Python 便携版，macOS 目前仅为源码预览。工具支持 **PDF、Word（.docx）与 PowerPoint（.ppt/.pptx）批量转换**，文件只在本机处理，不会上传到作者服务器。

![平台](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)
![Python](https://img.shields.io/badge/python-3.12-yellow)
![License](https://img.shields.io/badge/license-AGPL--3.0-blue)
![Release](https://img.shields.io/github/v/release/okeb910/pdf-word-converter)

---

## Windows v0.5.0 便携版（推荐）

**目标电脑不需要安装 Python，不需要运行 pip，也不需要下载源码。**

v0.5.0 只发布便携版。请从 [GitHub Releases](https://github.com/okeb910/pdf-word-converter/releases/tag/v0.5.0) 下载：

| 文件 | 说明 |
|------|------|
| `PDF-Word-PPT批量转换工具-v0.5.0-便携版-x64.exe` | **主程序**：下载后直接双击运行，无需安装 |
| `SHA256SUMS.txt` | 便携版 EXE 的 SHA-256 校验值 |
| `README-使用说明.txt` | 便携版使用、环境检测和校验说明 |
| `CHANGELOG.md` | v0.5.0 更新内容、修复和已知限制 |
| `THIRD-PARTY-LICENSES-v0.5.0.zip` | 第三方组件声明和完整许可文本 |

- 支持 **64 位 Windows 10 / 11**。
- EXE 暂未数字签名，SmartScreen 可能显示“未知发布者”；可点击“更多信息”后选择“仍要运行”。
- 便携版每次启动会先解压内置组件，再并行检测本机转换引擎；检测结果会逐项显示。
- 启动与环境检测日志位于 `%LOCALAPPDATA%\PDFWordConverter\logs`。

### 快速使用

1. 双击下载的便携版 EXE，无需安装。
2. 可点击添加按钮，也可把文件直接拖入转换队列。
3. PDF 可拖到“PDF → Word”或“PDF → PowerPoint”目标块；Word/PPT 拖入队列后自动转为 PDF。
4. 选择输出位置，默认输出到各源文件所在目录。
5. 点击“开始批量转换”。

同一批次只处理一种源格式；一次拖入多种源格式会整批拒绝并提示分开处理。程序会串行转换，单项失败不会中断剩余文件；取消会等待当前文件完成或失败后停止。

拖放要求程序与 Windows 文件资源管理器使用相同权限级别。通常直接双击运行即可，不要使用“以管理员身份运行”。

---

## macOS v0.5.0 源码预览

> **预览状态：目前只提供源码启动方式，没有 `.app` 或 `.dmg`，也尚未在真实 Mac 上完成启动、拖放和转换验收。** Windows v0.5.0 便携版不受此预览影响，仍是普通用户的稳定版本。

### 适用范围

- 目标系统：**macOS 13 或更高版本**。
- 目标架构：Apple Silicon `arm64` 与 Intel `x86_64`。
- 严格要求 **Python 3.12**；Python 3.11、3.13 或其他版本不会被启动器接受。
- 仅供源码预览和协助测试，不代表已经完成 macOS 正式发布质量验证。

### 源码启动

1. 下载或克隆完整仓库，保持文件目录结构不变。
2. 在 Finder 中双击 `启动工具.command`。
3. 首次运行会在项目目录创建 `.venv-macos`，并安装 `requirements-macos.txt` 中的固定版本依赖。
4. 后续启动会比较依赖文件的 SHA-256 并快速校验核心模块；依赖清单变化或环境损坏时才重新执行 pip 安装。

启动器会安全处理包含中文和空格的项目路径，不调用 Homebrew，也不会自动安装 Office 或 LibreOffice。找不到 Python 3.12 时，它会显示错误、打开 [Python 3.12.10 官方下载页](https://www.python.org/downloads/release/python-31210/) 后退出；请选择适用于 macOS 的 64 位 universal2 安装器。

如果 Finder 未保留脚本的可执行权限，可在终端执行一次：

```bash
cd "/完整路径/pdf-word-converter"
chmod +x "./启动工具.command"
"./启动工具.command"
```

### macOS 功能矩阵

| 方向 | macOS 源码预览引擎 | 当前说明 |
|------|---------------------|----------|
| **PDF → Word** | 内置整页图片模式（默认）；LibreOffice 备选 | 不提供 Microsoft Word 原生转换；LibreOffice PDF 导入兼容性有限 |
| **PDF → PowerPoint** | 内置 PyMuPDF + python-pptx 图片模式 | 不需要 Microsoft PowerPoint，元素不可单独编辑 |
| **Word → PDF** | Microsoft Word AppleScript 优先；LibreOffice 回退 | Office 路径需要 macOS Automation 权限 |
| **PPT/PPTX → PDF** | Microsoft PowerPoint AppleScript 优先；LibreOffice 回退 | Office 路径需要 macOS Automation 权限 |

以上矩阵描述 v0.5.0 的源码预览目标。仓库已经配置 `macos-15`（Apple Silicon）与 `macos-15-intel`（Intel）的 GitHub Actions 源码测试，每次推送和拉取请求的结果以仓库 Actions 页面为准；自动化测试不能替代真实 `arm64` 和 `x86_64` Mac 上的端到端验收。

### Automation 权限

首次通过源码启动器调用 Microsoft Word 或 PowerPoint 时，macOS 可能提示“终端”或“Python”要控制相应 Office 应用：

- 选择“允许”后，AppleScript 后端才能打开文档并导出 PDF。
- 如果曾选择拒绝，可前往“系统设置 → 隐私与安全性 → 自动化”，允许终端或 Python 控制 Microsoft Word / PowerPoint。
- Automation 权限不等于 Office 许可证；Word 和 PowerPoint 仍需正常安装、登录并完成授权。
- 不授予权限时，可安装 LibreOffice 并使用回退引擎；启动器不会替用户安装它。

### 预览限制

- 不提供签名、公证、Gatekeeper 验证过的 `.app` 或 `.dmg`。
- 不保证 Finder 拖放、Office Automation、LibreOffice 发现和转换效果已适配所有 Mac。
- 正式 macOS 发布前仍需分别在 Apple Silicon 与 Intel Mac 上完成启动、权限拒绝/允许、中文空格路径和全部方向的转换测试。

---

## Windows v0.5.0 转换能力

| 方向 | 默认方式 | 依赖与说明 |
|------|----------|------------|
| **PDF → Word** | Microsoft Word 原生 | 也可选择内置整页图片或 LibreOffice |
| **Word → PDF** | Word 优先 | Word 不可用时回退 LibreOffice |
| **PDF → PowerPoint** | 内置自适应高清图片 | 无需 Office；每页对应一张幻灯片，元素不可单独编辑 |
| **PPT/PPTX → PDF** | PowerPoint 优先 | PowerPoint 不可用时回退 LibreOffice |

### PDF → Word 三种方式

| 方式 | 文字可编辑 | 说明 |
|------|------------|------|
| **Microsoft Word 原生** | 通常可以 | 推荐，质量取决于 Word 和 PDF 结构 |
| **页面转高清图片嵌入** | 否 | 整页图片，版式较稳定 |
| **LibreOffice** | 视结果 | 免费备选，PDF 导入兼容性有限 |

### PDF → PowerPoint 高保真模式

- 使用 PyMuPDF 以约 200 DPI 自适应渲染，并限制异常大页面的像素尺寸。
- 每个 PDF 页面生成一张幻灯片。
- 使用第一页确定整份 PPTX 的画布比例。
- 横竖混合或尺寸不同的页面会等比居中并留白，不会裁切。
- 该模式以外观还原为目标，文字、表格和图形不会成为可单独编辑的 PowerPoint 元素。

### Windows 环境检测与安装提示

每次启动会并行检测内置组件、Microsoft Word、Microsoft PowerPoint、LibreOffice 和 winget，但**不会在启动时主动要求安装 Office**。

- Word、PowerPoint 和 LibreOffice 只做注册或路径浅检测，不会为了检测而启动外部程序。
- 已发现的外部引擎显示“已安装/使用时验证”，首次实际转换时才验证启动和导出能力。
- 当前转换方向所需的引擎检测完成后即可开始，不会被无关引擎阻塞。
- 用户选择不可用的 Word 引擎时，才询问是否安装 Microsoft Office。
- Word/PPT → PDF 没有可用引擎时，才询问是否安装 LibreOffice。
- 安装命令只有在用户确认后才执行；winget 不可用或安装失败时打开官方下载页。
- Office 仍需要有效许可证和 Microsoft 账号，程序不会提供许可证。

### 输出命名规则

程序不会覆盖已有文件，依次生成：

```text
原名.pptx/pdf/docx
原名_converted.pptx/pdf/docx
原名_converted_2.pptx/pdf/docx
...
```

---

## 开发者：从源码运行

### 环境要求

- Windows 源码与便携版构建：Python 3.12 x64。
- macOS 源码预览：macOS 13+、`arm64` 或 `x86_64`，且必须使用 Python 3.12。
- Microsoft Word、Microsoft PowerPoint 和 LibreOffice 按平台和转换方向选装。
- macOS 的 Word/PPT → PDF Office 路径需要系统 Automation 权限。

### Windows 源码启动

```bat
git clone https://github.com/okeb910/pdf-word-converter.git
cd pdf-word-converter
pip install -r requirements.txt
python pdf_word_converter.py
```

以上命令只适用于开发源码。普通用户应直接运行 Releases 中的便携版 EXE。

### macOS v0.5.0 源码预览启动

```bash
chmod +x "./启动工具.command"
"./启动工具.command"
```

不使用双击脚本时，也可以在终端执行标准命令：

```bash
cd "/完整路径/pdf-word-converter"
python3.12 -m venv .venv-macos
"./.venv-macos/bin/python" -m pip install --requirement requirements-macos.txt
"./.venv-macos/bin/python" launcher.py
```

`.command` 会严格检查 Python 3.12、维护项目内 `.venv-macos`，并根据 `requirements-macos.txt` 的 SHA-256 与核心模块导入结果决定是否安装或修复依赖。它不会调用 Homebrew。当前仓库没有 macOS `.app`/`.dmg` 构建或发布流程。

### 主要依赖

| 包 | 用途 |
|----|------|
| `PyMuPDF` | PDF 页面渲染 |
| `python-docx` | 写入和修补 DOCX |
| `python-pptx` | 生成高保真图片型 PPTX |
| `comtypes` | 仅 Windows：调用 Word 和 PowerPoint COM |
| `tkinterdnd2` | Windows/macOS 文件拖放 |

### 项目结构

```text
.
├── pdf_word_converter.py   # 转换引擎、后处理与 GUI
├── conversion_specs.py     # 转换方向和输出格式描述
├── drop_logic.py           # 拖放文件分类与混合格式校验
├── batch_logic.py          # 输出编号与串行批次调度
├── app_environment.py      # 冻结环境检测与按需安装
├── launcher.py             # 打包版启动检查与错误日志
├── build_release.ps1       # 本地构建脚本
├── 启动工具.command           # macOS Python 3.12 源码启动器
├── requirements-macos.txt  # macOS 源码预览固定依赖
├── packaging/              # PyInstaller、manifest、版本信息和第三方许可证
├── .github/workflows/      # Windows/macOS 源码测试，不构建发布物
├── tests/                  # 自动化测试
└── requirements*.txt       # 运行与构建依赖
```

### Windows 便携版构建与源码测试

Windows 构建机需要 Python 3.12 x64；脚本会在仓库本地创建 `.venv-build` 并安装固定版本的构建依赖：

```powershell
.\build_release.ps1 -PythonExe "C:\path\to\python.exe" -ReleaseDir "C:\path\to\release"
```

未指定 `-ReleaseDir` 时，v0.5.0 Windows 便携版默认生成到 `release\v0.5.0`。v0.5.0 已完成本地构建、启动和 Windows 转换验证；原 v0.4.1 Release、标签及 EXE 继续保留。

运行测试：

```bat
python -m unittest discover -s tests -v
```

macOS 启动器语法检查：

```bash
bash -n "./启动工具.command"
```

---

## 能力边界

1. PDF 的最终绘制结果通常没有原始 Word/PPT 的文本框、母版和图层语义，因此无法同时保证外观完全一致和元素完全可编辑。
2. 图片模式强调视觉还原，不适合直接修改文字。
3. 复杂多栏、扫描件、特殊字体和表单可能降低 Word/LibreOffice 转换质量。
4. 安全取消不会强杀当前 Word、PowerPoint 或 LibreOffice 进程。
5. 用于合同、公文、试卷和财务资料前，请人工核对输出结果。

## 许可证

本项目以 [GNU AGPL-3.0](LICENSE) 发布，并采用 PyMuPDF 的 AGPL 许可选项。Microsoft Office、LibreOffice 及其他第三方组件仍遵循各自许可条款，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。发布便携版时，对应源码应保留在同一 GitHub 仓库和版本标签中。
