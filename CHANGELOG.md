# 更新说明

## v0.5.0（2026-07-27）

### Windows 便携版

- 继续采用单文件便携版，目标电脑不需要安装 Python，也不会在启动时运行 pip。
- 启动检测改为并行浅检测：只检查 Office 注册和 LibreOffice 路径，不再为了检测而启动 Word、PowerPoint 或 LibreOffice。
- 本机测试中，外部引擎总检测时间由约 17 至 20 秒降至约 0.1 至 0.2 秒；不同电脑会因系统环境而有所差异。
- Word、PowerPoint 和 LibreOffice 改为首次实际转换时验证；原生 Office 转换失败时会清理半成品并自动回退 LibreOffice。
- 修复结构化引擎状态被错误当作“可用”的问题，避免缺失、启动失败或超时的引擎进入转换流程。
- 保留 PDF→Word、Word→PDF、PDF→PPTX、PPT/PPTX→PDF、批量队列、拖放、自定义输出目录、防覆盖命名和安全取消。

### macOS 源码预览

- 同一套源码新增 macOS 13+、Apple Silicon 和 Intel 的平台适配。
- Word→PDF 与 PPT/PPTX→PDF 可通过 Microsoft Office AppleScript 执行，并在不可用时回退 LibreOffice。
- PDF→Word 默认使用内置整页图片模式；PDF→PPTX 继续使用内置高清图片模式。
- 新增 `启动工具.command` 和固定版本的 `requirements-macos.txt`，首次运行创建项目内虚拟环境，依赖未变化时不重复安装。
- TkDND 加载失败时退化为普通 Tk 窗口，不阻止通过按钮选择文件。
- macOS Office 转换只关闭本工具打开的文档，不退出用户原本运行的 Word 或 PowerPoint。
- 新增 Apple Silicon 与 Intel 的 GitHub Actions 源码测试配置。

### 可靠性与发布

- 新增跨平台 `PlatformServices`、`ConversionBackend` 和 `EngineStatus`，统一平台行为和错误状态。
- Windows 与 macOS 的 Office/LibreOffice 回退前都会删除当前转换产生的不完整目标文件。
- 改善 Word COM 异常清理，确保文档、Office 实例和 COM 引用在失败路径中释放。
- 日志继续写入用户目录；启动错误使用可读的图形提示，不依赖控制台窗口。
- Windows 便携版版本资源、manifest、依赖清单和第三方许可证更新为 v0.5.0。
- 发布包包含 AGPL-3.0 主许可证、第三方组件声明和完整第三方许可文本。

### 验证结果

- Windows 自动化测试：105 项通过，1 项仅因测试机未开放符号链接权限而跳过。
- Windows 实际验证：Word PDF↔DOCX、PowerPoint PPTX→PDF、LibreOffice DOCX/PPTX→PDF，以及中文和空格路径。
- Windows 便携版已验证 x64 架构、无控制台窗口、中文界面、正常关闭、无残留进程和 SHA-256 一致性。
- Windows Defender 对最终便携版扫描未发现威胁。

### 已知限制

- macOS 当前仍是源码预览，没有 `.app` 或 `.dmg`，也尚未在真实 Mac 上完成 Office Automation、Finder 拖放和端到端转换验收。
- Windows 便携版尚未数字签名，SmartScreen 可能显示“未知发布者”。
- PDF 图片模式以外观还原为目标，生成的 Word/PPT 内容不能逐项编辑。
- LibreOffice 的 PDF→Word 兼容性有限，复杂文件需要人工核对。
- Microsoft Office 转换仍需要用户自己的有效许可证和账号。

## v0.4.1

- Windows 单便携版发布，加入拖放目标区和并行环境检测。
- 修正便携版中文文件名、说明文件编码、许可证和 GitHub 发布内容。
