# PDF ↔ Word/PPT 批量转换工具

Windows 本地 **PDF、Word（.docx）与 PowerPoint（.ppt/.pptx）批量转换** 桌面工具。文件只在本机处理，不会上传到作者服务器。

![平台](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-yellow)
![License](https://img.shields.io/badge/license-MIT-green)
![Release](https://img.shields.io/github/v/release/okeb910/pdf-word-converter)

---

## 普通用户：下载单便携版（推荐）

**目标电脑不需要安装 Python，不需要运行 pip，也不需要下载源码。**

v0.4.0 只发布便携版。请从 [GitHub Releases](https://github.com/okeb910/pdf-word-converter/releases/tag/v0.4.0) 下载：

| 文件 | 说明 |
|------|------|
| `PDF-Word-PPT批量转换工具-v0.4.0-便携版-x64.exe` | **主程序**：下载后直接双击运行，无需安装 |
| `SHA256SUMS.txt` | 便携版 EXE 的 SHA-256 校验值 |

- 支持 **64 位 Windows 10 / 11**。
- EXE 暂未数字签名，SmartScreen 可能显示“未知发布者”；可点击“更多信息”后选择“仍要运行”。
- 便携版每次启动会先解压内置组件并在后台检测环境，可能需要等待十几秒，请勿连续重复双击。
- 启动与环境检测日志位于 `%LOCALAPPDATA%\PDFWordConverter\logs`。

### 快速使用

1. 双击下载的便携版 EXE，无需安装。
2. 点击“添加 PDF”“添加 Word”或“添加 PowerPoint”，可一次多选同类文件。
3. 添加 PDF 时，在“转换目标”中选择 Word 或 PowerPoint。
4. 选择输出位置，默认输出到各源文件所在目录。
5. 点击“开始批量转换”。

同一批次只处理一种源格式，程序会串行转换。单项失败不会中断剩余文件；取消会等待当前文件完成或失败后停止。

---

## 转换能力

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

### 环境检测与安装提示

每次启动会在后台检测内置组件、Microsoft Word、Microsoft PowerPoint、LibreOffice 和 winget，但**不会在启动时主动要求安装 Office**。

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

- Windows
- Python 3.8+，构建发布物建议使用 Python 3.12 x64
- Microsoft Word、Microsoft PowerPoint 和 LibreOffice 按转换方向选装

### 安装与启动

```bat
git clone https://github.com/okeb910/pdf-word-converter.git
cd pdf-word-converter
pip install -r requirements.txt
python pdf_word_converter.py
```

以上命令只适用于开发源码。普通用户应直接运行 Releases 中的便携版 EXE。

### 主要依赖

| 包 | 用途 |
|----|------|
| `PyMuPDF` | PDF 页面渲染 |
| `python-docx` | 写入和修补 DOCX |
| `python-pptx` | 生成高保真图片型 PPTX |
| `comtypes` | 调用 Word 和 PowerPoint COM |

### 项目结构

```text
.
├── pdf_word_converter.py   # 转换引擎、后处理与 GUI
├── conversion_specs.py     # 转换方向和输出格式描述
├── batch_logic.py          # 输出编号与串行批次调度
├── app_environment.py      # 冻结环境检测与按需安装
├── launcher.py             # 打包版启动检查与错误日志
├── build_release.ps1       # 本地构建脚本
├── packaging/              # PyInstaller、manifest、Inno Setup
├── tests/                  # 自动化测试
└── requirements*.txt       # 运行与构建依赖
```

### 构建与测试

构建机需要 Inno Setup 6：

```powershell
.\build_release.ps1 -PythonExe "C:\path\to\python.exe" -ReleaseDir "C:\path\to\release"
```

运行测试：

```bat
python -m unittest discover -s tests -v
```

---

## 能力边界

1. PDF 的最终绘制结果通常没有原始 Word/PPT 的文本框、母版和图层语义，因此无法同时保证外观完全一致和元素完全可编辑。
2. 图片模式强调视觉还原，不适合直接修改文字。
3. 复杂多栏、扫描件、特殊字体和表单可能降低 Word/LibreOffice 转换质量。
4. 安全取消不会强杀当前 Word、PowerPoint 或 LibreOffice 进程。
5. 用于合同、公文、试卷和财务资料前，请人工核对输出结果。

## 许可证

本项目以 [MIT License](LICENSE) 发布。Microsoft Office、LibreOffice、PyMuPDF 等第三方组件遵循各自许可条款。
