PDF Word PPT Converter v0.4.0 使用说明
=======================================

系统要求
--------
- 64 位 Windows 10 或 Windows 11。
- 目标电脑不需要安装 Python，也不会在启动时运行 pip。
- 便携版和安装版暂未数字签名，Windows SmartScreen 可能显示“未知发布者”。

启动方式
--------
- 便携版：双击 PDF-Word-PPT批量转换工具-v0.4.0-便携版-x64.exe。
- 安装版：运行 PDF-Word-PPT批量转换工具-v0.4.0-安装版-x64.exe。
- 启动时只在后台检测环境，不会主动要求安装 Office。

转换方向
--------
- PDF → Word：可选 Word 原生、内置图片或 LibreOffice。
- Word → PDF：优先 Word，不可用时使用 LibreOffice。
- PDF → PowerPoint：内置自适应高清图片模式，每页生成一张幻灯片。
- PPT/PPTX → PDF：优先 PowerPoint，不可用时使用 LibreOffice。

PDF → PowerPoint 说明
---------------------
- 不需要安装 Microsoft PowerPoint。
- 以原页面外观还原为目标，横竖混合页面等比居中且不裁切。
- 幻灯片中的内容是整页图片，文字、表格和图形不能单独编辑。

按需安装提示
------------
- 只有选择不可用的 Word 引擎时才询问安装 Microsoft Office。
- Word/PPT → PDF 没有可用引擎时才询问安装 LibreOffice。
- 程序只有在用户确认后才执行安装命令。
- Office 需要许可证和 Microsoft 账号，程序不会提供许可证。

批量转换与输出
--------------
- 同一批次只添加 PDF、DOCX 或 PowerPoint 文件，程序会串行处理。
- 默认输出到各源文件目录，也可选择统一输出目录。
- 不覆盖已有文件，依次使用：原名、原名_converted、原名_converted_2。
- 取消会等待当前文件完成或失败，不会强制结束外部进程。

日志
----
启动和环境检测日志位于：
%LOCALAPPDATA%\PDFWordConverter\logs

校验
----
SHA256SUMS.txt 包含两个 EXE 的 SHA-256，可用 PowerShell 核对：
Get-FileHash '.\PDF-Word-PPT批量转换工具-v0.4.0-便携版-x64.exe' -Algorithm SHA256
Get-FileHash '.\PDF-Word-PPT批量转换工具-v0.4.0-安装版-x64.exe' -Algorithm SHA256
