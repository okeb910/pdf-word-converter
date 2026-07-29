# PDF ↔ Word/PPT 批量转换工具

当前版本为 **v0.5.1**。工具支持 PDF、Word（`.docx`）和 PowerPoint（`.ppt/.pptx`）批量转换，文件只在本机处理，不会上传到作者服务器。

![平台](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)
![Python](https://img.shields.io/badge/python-3.12-yellow)
![License](https://img.shields.io/badge/license-AGPL--3.0-blue)
![Release](https://img.shields.io/github/v/release/okeb910/pdf-word-converter)

## 主要能力

| 转换方向 | 默认方式 | 结果说明 |
|------|------|------|
| **PDF → Word** | Microsoft Word 可编辑模式 | 支持复杂表格检查；检测结果不确定时继续转换并提醒人工核对 |
| **Word → PDF** | Microsoft Word | Word 不可用时可回退 LibreOffice |
| **PDF → PowerPoint** | 内置高清图片模式 | 每页生成一张幻灯片，保持页面外观，内容不可单独编辑 |
| **PPT/PPTX → PDF** | Microsoft PowerPoint | PowerPoint 不可用时可回退 LibreOffice |

- 支持多选、拖放、串行批量转换和自定义输出目录。
- 自动避免覆盖已有文件，单项失败不会中断剩余文件。
- 支持安全取消；外部 Office 已开始处理时会等待当前文件结束。
- 启动时并行检测 Word、PowerPoint、LibreOffice 和内置组件，不会主动要求安装 Office。

## Windows 便携版

支持 **64 位 Windows 10 / 11**。目标电脑不需要安装 Python，也不需要运行 `pip`。

从 [GitHub Releases v0.5.1](https://github.com/okeb910/pdf-word-converter/releases/tag/v0.5.1) 下载：

| 文件 | 说明 |
|------|------|
| `PDF-Word-PPT-Converter-v0.5.1-Portable-x64.exe` | 主程序，下载后直接双击运行 |
| `SHA256SUMS.txt` | EXE 的 SHA-256 校验值 |
| `README-v0.5.1-zh-CN.txt` | 中文使用说明 |
| `THIRD-PARTY-LICENSES-v0.5.1.zip` | 第三方许可文本 |

EXE 暂未数字签名，SmartScreen 可能显示“未知发布者”。可点击“更多信息”，确认文件来自本仓库后选择“仍要运行”。

## 使用方法

1. 双击便携版 EXE。
2. 点击添加按钮，或把 PDF、DOCX、PPT、PPTX 拖入窗口。
3. PDF 选择输出 Word 或 PowerPoint；Word/PPT 自动转换为 PDF。
4. 选择输出位置，默认使用源文件所在目录。
5. 点击“开始批量转换”。
6. 转换完成后查看每项状态；出现“转换质量提醒”时请对照原文件核对。

同一批次只处理一种源格式。拖放要求程序与资源管理器使用相同权限级别，通常不要使用“以管理员身份运行”。

### PDF → Word 模式

| 模式 | 适用内容 | 说明 |
|------|------|------|
| **Word 可编辑模式** | 含文字、表格、方框或勾选的 PDF | 优先保留可编辑内容；复杂版式可能重排 |
| **整页图片模式** | 扫描件或纯图形 PDF | 版式较稳定，但文字和表格不可编辑，不提供 OCR |
| **LibreOffice 兼容模式** | 没有 Word 时 | 免费备选，PDF 导入兼容性有限 |

程序会检查可选择文字、表格迹象、方框、勾选和异常字体。无法可靠判断时，会优先生成可编辑结果并显示质量提醒，不会把含可选择文字的文件静默转换成整页图片。重要表格、合同、试卷和财务资料转换后仍需人工核对。

### 输出命名

程序不会覆盖已有文件，依次生成：

```text
原名.docx/pdf/pptx
原名_converted.docx/pdf/pptx
原名_converted_2.docx/pdf/pptx
```

## 环境与引擎

- Word、PowerPoint 和 LibreOffice 在首次实际转换时验证，不会为了启动检测而打开外部程序。
- 只有所选方向缺少可用引擎时才显示安装提示；执行安装前一定会询问用户。
- Microsoft Office 需要有效许可证和 Microsoft 账号，程序不会提供许可证。
- 启动与环境检测日志位于 `%LOCALAPPDATA%\PDFWordConverter\logs`。

## macOS 源码预览

macOS 版本目前只提供源码启动方式，没有 `.app` 或 `.dmg`，也尚未完成真实 Mac 端到端验收。

- 系统：macOS 13 或更高版本。
- 架构：Apple Silicon `arm64` 与 Intel `x86_64`。
- 环境：Python 3.12。

在 Finder 中双击 `启动工具.command`，或在终端运行：

```bash
chmod +x "./启动工具.command"
"./启动工具.command"
```

macOS 支持 PDF → Word（LibreOffice 可编辑兼容模式；扫描件可用图片模式）、PDF → PPTX、Word → PDF 和 PPT/PPTX → PDF。调用 Microsoft Word 或 PowerPoint 时，需要在“系统设置 → 隐私与安全性 → 自动化”中允许终端或 Python 控制对应应用。

## 源码运行

Windows 开发环境需要 Python 3.12 x64：

```bat
pip install -r requirements.txt
python pdf_word_converter.py
```

详细版本变化和测试记录见 [CHANGELOG.md](CHANGELOG.md)。

## 使用限制

- PDF 通常不包含原始 Word/PPT 的完整表格、字体、图层和分页语义，无法保证所有文件完全还原。
- PDF → PPTX 以页面外观为目标，幻灯片元素不可单独编辑。
- LibreOffice 是兼容性回退方案，效果可能低于 Microsoft Office。
- 安全取消不会强制结束 Word、PowerPoint 或 LibreOffice。

## 许可证

本项目以 [GNU AGPL-3.0](LICENSE) 发布，并采用 PyMuPDF 的 AGPL 许可选项。第三方组件许可见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
