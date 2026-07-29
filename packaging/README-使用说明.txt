PDF ↔ Word/PPT 批量转换工具 v0.5.1 Windows 便携版
===================================================

系统要求
--------
- 64 位 Windows 10 或 Windows 11。
- 不需要安装 Python，也不会在启动时运行 pip。
- EXE 暂未数字签名，SmartScreen 可能显示“未知发布者”。

转换能力
--------
- PDF → Word：优先使用 Microsoft Word 生成可编辑内容；也支持扫描件整页图片模式和 LibreOffice 兼容模式。
- Word → PDF：优先使用 Microsoft Word，不可用时回退 LibreOffice。
- PDF → PowerPoint：每个 PDF 页面生成一张高清图片型幻灯片，保持页面外观，内容不可单独编辑。
- PPT/PPTX → PDF：优先使用 Microsoft PowerPoint，不可用时回退 LibreOffice。
- 支持批量队列、文件拖放、自定义输出目录、防覆盖命名和安全取消。

使用方法
--------
1. 双击 PDF-Word-PPT-Converter-v0.5.1-Portable-x64.exe。
2. 点击添加按钮，或把 PDF、DOCX、PPT、PPTX 拖入窗口。
3. PDF 选择输出 Word 或 PowerPoint；Word/PPT 自动转换为 PDF。
4. 选择输出位置，默认使用源文件所在目录。
5. 点击“开始批量转换”。
6. 转换完成后查看每项状态；出现质量提醒时请对照原文件核对。

同一批次只处理一种源格式。通常不要使用“以管理员身份运行”，否则资源管理器可能无法拖放文件到窗口。

PDF → Word 说明
----------------
- 含可选择文字、表格、方框或勾选的 PDF 会优先使用可编辑引擎。
- 扫描件或纯图形 PDF 可使用整页图片模式；该模式不提供 OCR，文字和表格不可编辑。
- 无法可靠判断复杂版式时，程序会继续生成可编辑结果并显示质量提醒。
- 程序不会把含可选择文字的 PDF 静默转换成整页图片。
- PDF 通常不保存完整的原始表格、字体和分页语义，重要文件转换后必须人工核对。

环境检测
--------
- 启动时并行检测 Word、PowerPoint、LibreOffice、内置组件和 winget。
- 启动检测不会打开 Office，也不会主动要求安装 Office。
- 外部引擎在首次实际转换时验证；只有所选方向缺少可用引擎时才显示安装提示。
- Office 需要有效许可证和 Microsoft 账号，程序不会提供许可证。

批量转换与输出
--------------
- 默认输出到源文件目录，也可选择统一输出目录。
- 不覆盖已有文件，依次使用原名、原名_converted、原名_converted_2。
- 单项失败不会中断剩余文件。
- 点击取消后，程序不会强制结束 Office 或 LibreOffice；当前文件结束后停止。

日志
----
%LOCALAPPDATA%\PDFWordConverter\logs

文件校验
--------
EXE SHA-256：
a42e0971722fe73e0f72612803d343f5ba73916702bf112d648a99a355393192

PowerShell 校验命令：
Get-FileHash '.\PDF-Word-PPT-Converter-v0.5.1-Portable-x64.exe' -Algorithm SHA256

许可证
------
本项目以 GNU AGPL-3.0 发布，并采用 PyMuPDF 的 AGPL 许可选项。
第三方组件许可见 THIRD_PARTY_NOTICES.md；对应源码位于本仓库 v0.5.1 标签。
