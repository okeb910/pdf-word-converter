PDF Word Converter v0.3.0 使用说明
===================================

系统要求
--------
- 64 位 Windows 10 或 Windows 11。
- 目标电脑不需要安装 Python，也不会在启动时运行 pip。
- 便携版和安装版暂未数字签名，Windows SmartScreen 可能显示“未知发布者”。

启动方式
--------
- 便携版：直接双击 PDF-Word批量转换工具-v0.3.0-便携版-x64.exe。
- 安装版：运行 PDF-Word批量转换工具-v0.3.0-安装版-x64.exe，完成后从开始菜单启动。
- 每次启动都会在后台检查内置组件、Microsoft Word、LibreOffice 和 winget。

转换引擎
--------
- PDF → Word 推荐使用 Microsoft Word 原生转换。
- “页面转高清图片嵌入”无需 Office，版式较稳定，但文字不可编辑。
- LibreOffice PDF → Word 兼容性有限；导出失败时可改用图片模式重试。
- Word → PDF 会优先使用 Microsoft Word，没有 Word 时使用 LibreOffice。

按需安装提示
------------
- 未检测到 Word 时，每次启动最多询问一次是否通过 winget 安装 Microsoft Office。
- Office 需要许可证和 Microsoft 账号，安装可能要求管理员权限。
- LibreOffice 仅在选择其引擎或 Word → PDF 没有可用引擎时询问安装。
- 程序只有在用户确认后才会执行安装命令；winget 不可用或安装失败时会打开官方页面。

批量转换与输出
--------------
- 同一批次只添加 PDF 或只添加 DOCX，程序会串行处理。
- 默认输出到各源文件目录，也可以选择统一输出目录。
- 不覆盖已有文件，依次使用：原名、原名_converted、原名_converted_2。
- 取消会等待当前文件完成或失败，不会强制结束 Word 或 LibreOffice。

日志
----
启动和环境检测日志位于：
%LOCALAPPDATA%\PDFWordConverter\logs

校验
----
SHA256SUMS.txt 包含两个 EXE 的 SHA-256，可用 PowerShell 核对：
Get-FileHash '.\PDF-Word批量转换工具-v0.3.0-便携版-x64.exe' -Algorithm SHA256
Get-FileHash '.\PDF-Word批量转换工具-v0.3.0-安装版-x64.exe' -Algorithm SHA256
