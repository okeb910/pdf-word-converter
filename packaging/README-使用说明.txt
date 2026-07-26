PDF ↔ Word/PPT 批量转换工具 v0.4.1 使用说明
==========================================

系统要求
--------
- 64 位 Windows 10 或 Windows 11。
- 目标电脑不需要安装 Python，也不会在启动时运行 pip。
- 便携版暂未数字签名，Windows SmartScreen 可能显示“未知发布者”。

启动与添加文件
--------------
- 双击 PDF-Word-PPT批量转换工具-v0.4.1-便携版-x64.exe。
- 可点击添加按钮，也可把 PDF、DOCX、PPT 或 PPTX 拖入转换队列。
- PDF 可直接拖到“PDF → Word”或“PDF → PowerPoint”目标块。
- Word/PPT 拖入普通队列后自动转换为 PDF。
- 一次拖入多种源格式会整批拒绝，请分开处理。
- 请直接双击运行，不要使用“以管理员身份运行”，否则 Windows 可能阻止资源管理器拖放。

转换方向
--------
- PDF → Word：可选 Word 原生、内置图片或 LibreOffice。
- Word → PDF：优先 Word，不可用时使用 LibreOffice。
- PDF → PowerPoint：内置自适应高清图片模式，每页生成一张幻灯片。
- PPT/PPTX → PDF：优先 PowerPoint，不可用时使用 LibreOffice。

环境检测
--------
- Word、PowerPoint、LibreOffice、内置组件和 winget 会并行检测并逐项显示。
- 当前转换方向所需检测完成后即可操作，不等待无关引擎。
- 深度检测会实际启动并关闭本机 Office/LibreOffice，耗时取决于电脑性能。
- 启动检测不会主动要求安装 Office；只有真正缺少所选引擎时才提示。
- Office 需要许可证和 Microsoft 账号，程序不会提供许可证。

PDF → PowerPoint 说明
---------------------
- 不需要安装 Microsoft PowerPoint。
- 以原页面外观还原为目标，横竖混合页面等比居中且不裁切。
- 幻灯片中的内容是整页图片，文字、表格和图形不能单独编辑。

批量转换与输出
--------------
- 同一批次只处理一种源格式，程序会串行处理。
- 默认输出到各源文件目录，也可选择统一输出目录。
- 不覆盖已有文件，依次使用：原名、原名_converted、原名_converted_2。
- 取消会等待当前文件完成或失败，不会强制结束外部进程。

日志
----
启动和环境检测日志位于：
%LOCALAPPDATA%\PDFWordConverter\logs

校验
----
SHA256SUMS.txt 包含便携版 EXE 的 SHA-256，可用 PowerShell 核对：
Get-FileHash '.\PDF-Word-PPT批量转换工具-v0.4.1-便携版-x64.exe' -Algorithm SHA256
