# PDF ↔ Word/PPT 批量转换工具

当前版本为 **v0.5.1**，由同一套源码支持 Windows 与 macOS；Windows 面向普通用户提供免 Python 便携版，macOS 目前仍为源码预览。v0.5.1 新增短文档可编辑表格硬门禁、长文档快速预检、表格拓扑与边框修复，以及转换完成后的质量提醒。工具支持 **PDF、Word（.docx）与 PowerPoint（.ppt/.pptx）批量转换**，文件只在本机处理，不会上传到作者服务器。

![平台](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)
![Python](https://img.shields.io/badge/python-3.12-yellow)
![License](https://img.shields.io/badge/license-AGPL--3.0-blue)
![Release](https://img.shields.io/github/v/release/okeb910/pdf-word-converter)

---

## Windows v0.5.1 便携版（推荐）

**目标电脑不需要安装 Python，不需要运行 pip，也不需要下载源码。**

v0.5.1 只发布便携版。请从 [GitHub Releases](https://github.com/okeb910/pdf-word-converter/releases/tag/v0.5.1) 下载：

| 文件 | 说明 |
|------|------|
| `PDF-Word-PPT-Converter-v0.5.1-Portable-x64.exe` | **主程序**：下载后直接双击运行，无需安装 |
| `SHA256SUMS.txt` | 便携版 EXE 的 SHA-256 校验值 |
| `README-使用说明.txt` | 便携版使用、环境检测和校验说明 |
| `CHANGELOG.md` | v0.5.1 更新内容、修复和已知限制 |
| `THIRD-PARTY-LICENSES-v0.5.1.zip` | 第三方组件声明和完整许可文本 |

- 支持 **64 位 Windows 10 / 11**。
GitHub 会自动清理含中文的附件文件名，因此下载附件统一使用可稳定保留的英文文件名；Release 页面仍以中文标签标明“便携版”和各说明文件用途。

- EXE 暂未数字签名，SmartScreen 可能显示“未知发布者”；可点击“更多信息”后选择“仍要运行”。
- 便携版每次启动会先解压内置组件，再并行检测本机转换引擎；检测结果会逐项显示。
- 启动与环境检测日志位于 `%LOCALAPPDATA%\PDFWordConverter\logs`。

### 快速使用

1. 双击下载的便携版 EXE，无需安装。
2. 可点击添加按钮，也可把文件直接拖入转换队列。
3. PDF 可拖到“PDF → Word”或“PDF → PowerPoint”目标块；Word/PPT 拖入队列后自动转为 PDF。
4. 选择输出位置，默认输出到各源文件所在目录。
5. 点击“开始批量转换”，程序会串行处理队列并逐项显示结果。

同一批次只处理一种源格式；一次拖入多种源格式会整批拒绝并提示分开处理。程序会串行转换，单项失败不会中断剩余文件；取消会等待当前文件完成或失败后停止。

> 本节描述 GitHub Release 中的 v0.5.1 Windows 便携版。

拖放要求程序与 Windows 文件资源管理器使用相同权限级别。通常直接双击运行即可，不要使用“以管理员身份运行”。

---

## macOS v0.5.1 源码预览

> **预览状态：目前只提供 v0.5.1 源码启动方式，没有 `.app` 或 `.dmg`，也尚未在真实 Mac 上完成启动、拖放和转换验收。** Windows v0.5.1 便携版不受此限制。

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
| **PDF → Word** | LibreOffice 可编辑兼容模式；内置图片仅限扫描件/纯图形 | 不提供 Microsoft Word 原生转换；含任何可选择文字时禁止图片化，没有 LibreOffice 则该项失败 |
| **PDF → PowerPoint** | 内置 PyMuPDF + python-pptx 图片模式 | 不需要 Microsoft PowerPoint，元素不可单独编辑 |
| **Word → PDF** | Microsoft Word AppleScript 优先；LibreOffice 回退 | Office 路径需要 macOS Automation 权限 |
| **PPT/PPTX → PDF** | Microsoft PowerPoint AppleScript 优先；LibreOffice 回退 | Office 路径需要 macOS Automation 权限 |

以上矩阵描述 v0.5.1 的源码预览目标。仓库已经配置 `macos-15`（Apple Silicon）与 `macos-15-intel`（Intel）的 GitHub Actions 源码测试，每次推送和拉取请求的结果以仓库 Actions 页面为准；自动化测试不能替代真实 `arm64` 和 `x86_64` Mac 上的端到端验收。

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

## Windows v0.5.1 转换能力

| 方向 | 默认方式 | 依赖与说明 |
|------|----------|------------|
| **PDF → Word** | Microsoft Word 原生 + 后台逐文件预检 | 只要含任何可选择文字就锁定可编辑引擎并校验 DOCX；图片模式仅在用户明确选择或确认、且预检成功证明没有可选择文字时使用 |
| **Word → PDF** | Word 优先 | Word 不可用时回退 LibreOffice |
| **PDF → PowerPoint** | 内置自适应高清图片 | 无需 Office；每页对应一张幻灯片，元素不可单独编辑 |
| **PPT/PPTX → PDF** | PowerPoint 优先 | PowerPoint 不可用时回退 LibreOffice |

### PDF → Word 三种方式

| 方式 | 文字可编辑 | 说明 |
|------|------------|------|
| **可编辑优先：Microsoft Word 原生** | 通常可以 | 保留 Word 的原始转换结果；复杂表格仍可能重排或分页 |
| **仅扫描件/纯图形：整页高清图片** | 否 | 默认约 300 DPI；不执行 OCR；只在用户明确选择或确认后使用，检测到任何可选择文字时会改用可编辑引擎 |
| **兼容模式：LibreOffice** | 视结果 | 免费备选，PDF 导入兼容性有限 |

### 复杂表格、方框与勾选

- v0.5.1 在批次工作线程逐文件检查全部页面的可选择文字、字体映射、PDF Widget、复选/勾选符号、常见符号字体、矢量路径和表格迹象。20 页以内的文档继续逐页提取线框表格、无边框表格、单元格拓扑和四边证据；超过 20 页的长文档改用全页快速文字与风险扫描，不再对每页重复执行多轮高成本表格重建。
- 另有一条不依赖主表格提取器的保守检查，直接寻找重复对齐的文字列和横纵线网格。主检测未确认表格、长文档未执行逐格表格基线、出现 U+FFFD 替换字符/私用区字符，或勾选状态无法自动证明时，只要有可编辑引擎就继续转换，并在批次完成时显示“转换质量提醒”，要求对照原 PDF 核对，不再因为检测不确定而直接不给输出。
- 只要 PDF 含任何可选择文字，不论是否已经明确识别为表格，都只允许 Word/LibreOffice 可编辑引擎；即使用户手动选择图片模式也会自动改用可编辑引擎，LibreOffice 失败后同样禁止回退整页图片。
- 每个文件独立决定引擎。整页图片不会被静默选中，也不会作为失败后的自动兜底；只有用户已经明确选择图片模式，或在提示中明确确认，并且预检成功证明完全没有可选择文字时才会使用。LibreOffice 仅在导出过滤器失败、没有可编辑内容保护且用户再次确认时才允许改用图片。
- PDF 预检发生异常时采用失败关闭：该项直接失败并继续批次，不会未经检查进入转换，也不会生成图片替代文件。预检结果同时绑定源文件身份，排队期间被替换或改写的文件必须重新加入。
- 检测出表格后，程序先记录源 PDF 的精确行列矩阵、每个单元格的行列跨度、列宽几何，以及每个逻辑单元格 top/left/bottom/right 四条边的可见描边证据。Word 如果把一个源单元格拆成多个相邻格，只在几何边界、逐格文字、样式和外边框能够唯一证明归属时进行事务式合并；拓扑已经正确但边框缺失或写成 `nil` 时仍会按源证据恢复。真实合并格内部不会误加线，源 PDF 确认无边框的表格也不会被强制套用网格。
- 对 20 页以内且源表格基线完整的文档，修复后继续执行硬门禁：DOCX 必须能正常打开并包含真实的 `w:tbl`、`w:tr`、`w:tc`；表格数量和顺序、有效行数、网格列数、独立单元格数、每个单元格的 `rowSpan`/`gridSpan` 拓扑、规范化文字和四边状态都必须与源 PDF 证据一致。总表格文字字符数也必须完全相等；少边、多边、少字或文字错格都会拒绝输出。空白表格只能凭完整拓扑通过。v0.5.1 不再使用“保留 80%”之类的宽松阈值。
- 门禁同时拒绝照片加伪表格、全页大图、文字错格或换格、合并关系变化、隐藏文字、近白文字、透明文字、异常小字号、不可证明可见的表格、文档保护、写保护以及表格内外的锁定内容控件。
- 每项转换先写入同目录的私有 `.partial.docx`。源基线可靠时只有结构修复和硬门禁全部通过后才发布正式文件；源结构或字体映射本身无法可靠验证时，只要 DOCX 可正常打开就保留可编辑结果，并在完成弹窗列出严格校验差异。真正的转换异常、损坏 DOCX、可靠基线下的确定漏字/错格，以及源文件在转换期间被改写仍会失败并清理不完整文件。
- 可编辑表格同时含 Widget、勾选、疑似矢量勾选或 Symbol/Wingdings 字体时，目前无法逐格证明状态不变，因此继续使用可编辑引擎并在完成后提醒人工核对，不提供静默图片回退；没有可用的 Word/LibreOffice 时仍会明确失败。
- 可编辑路径不再执行旧版的列宽、字号、上下标和标点猜测性后处理。拓扑完全通过门禁后，程序只对单个近满页大型表格执行受限版式修复：清理系统性重复的文字方框；仅当全部行都有有效显式高度时使用固定行高；仅对符合条件的表格收紧页边距和单元格上下内边距。孤立的合法文字边框和无法证明安全的格式不会被改写。
- 图片模式使用 RGB PNG、关闭 Office 自动图片压缩并限制长边和总像素；极端页面即使 1 DPI 仍超限时会明确报错，不会冒险分配超大内存。

#### 真实成绩表回归

- 本轮真实样本是可选择文字加矢量线框的成绩表，不是扫描图片。正确结构为 `33×26`、515 个逻辑单元格、516 个原始 `<w:tc>` 和 1075 个表格文字字符；唯一真实纵向合并是底部跨两行的“其中包括”单元格。
- 已修复相邻行 2-3 pt 边界偏移被误扩展为幽灵列、Word 错误纵向聚合，以及拓扑恢复后继承 `nil` 造成的缺线。源 PDF 的 515 格共有 2060 条逻辑外边；旧输出只有 1905 条，现已全部恢复。修复后行列跨度、逐格文字和逐格四边状态与源 PDF 证据完全一致，输出仍是可编辑 Word 表格，含 0 张图片。
- 修复后的 DOCX 可由 Microsoft Word 重新打开并导出为单页 A4 横向 PDF；放大检查三个课程区块和底部合并格，未发现漏线、漏字、裁切、黑块或越界文字。

#### 225 页年度报告回归

- 使用 225 页、168,956 个可选择非空字符的真实年度报告验证长文档路径。旧版逐页多轮表格提取耗时约 85.8 秒，快速预检仍扫描全部页面文字和风险标记，耗时约 0.6 秒。
- Microsoft Word 原生转换约 26.6 秒并生成可打开、可编辑的 DOCX。Word 输出含 268 个真实表格和 162,655 个可见非空字符；与源 PDF 相差 6,301 个字符，重新导出后由 225 页重排为 259 页，因此程序保留结果但必须弹出质量提醒，不能宣称严格还原。
- 导出的 259 页 PDF 已全部渲染检查：无黑页、无内容触边裁切，但有 9 页仅剩页码或极少内容。弹窗会要求重点核对文字、表格行列、边框、分页和勾选状态。

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

### macOS v0.5.1 源码预览启动

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
| `python-docx` | 生成图片型 DOCX，并验证 DOCX 包可正常打开 |
| `lxml` | 保留并校验 Word OOXML 命名空间，安全修复表格拓扑 |
| `python-pptx` | 生成高保真图片型 PPTX |
| `comtypes` | 仅 Windows：调用 Word 和 PowerPoint COM |
| `tkinterdnd2` | Windows/macOS 文件拖放 |

### 项目结构

```text
.
├── pdf_word_converter.py   # 转换引擎与 GUI
├── pdf_fidelity.py         # PDF 复杂版式预检
├── docx_table_repair.py    # 可证明时修复 Word 表格拆格与合并拓扑
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

未指定 `-ReleaseDir` 时，v0.5.1 Windows 便携版默认生成到 `release\v0.5.1`。构建脚本会先导入检查 `docx_table_repair` 与 `pdf_fidelity`，PyInstaller 配置也会显式收集这两个模块。当前完整自动化测试结果为 `289 passed, 1 skipped`。正式发布的 EXE 大小为 46,679,665 字节，SHA-256 为 `a42e0971722fe73e0f72612803d343f5ba73916702bf112d648a99a355393192`；已核对为 x64 Windows GUI、版本 `0.5.1.0`，中文空格路径约 5.13 秒出现窗口，环境检测约 133 ms，内置组件正常，Windows Defender 自定义扫描 0 检测。原 v0.5.0、v0.4.1 Release、标签及 EXE 继续保留。

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

1. PDF 的最终绘制结果通常没有原始 Word/PPT 的合并关系、文本框、母版和图层语义，因此无法同时保证外观完全一致和元素完全可编辑。
2. v0.5.1 同时使用主表格提取和独立的重复文字列/网格检查；只要 PDF 含任何可选择文字，就禁止整页图片和图片回退。疑似表格、可疑字体映射、勾选状态或长文档逐格基线无法证明时，会继续使用可编辑引擎并在完成后警告；预检本身异常、没有可编辑引擎或源文件被改写时仍失败。PDF 通常不保存原始表格语义，因此重要文件必须人工核对。
3. 对 20 页以内且源基线可靠的文档，DOCX 硬门禁要求源表与 Word 表格的数量、顺序、行列矩阵、独立单元格、行列跨度和逐格文字完全对应；保守修复只会处理能够由几何、文字和样式唯一证明的单元格。源基线本身不可靠或长文档采用快速预检时，程序优先保留可打开结果并显示具体差异，不等于结果已经通过严格保真验证。
4. 图片模式强调视觉还原，只适合用户明确选择的扫描件或纯图形内容，不适合直接修改文字，也不会被静默用作转换失败的兜底。
5. 复杂多栏、特殊字体和表单可能降低 Word/LibreOffice 转换质量。
6. 安全取消会在 PDF 预检的逐页和表格识别阶段协作停止，并每秒显示等待时间；如果 Word、PowerPoint 或 LibreOffice 已经开始处理当前文件，则等待该文件完成或失败后停止，不会强杀外部进程。
7. 用于合同、公文、试卷和财务资料前，请人工核对输出结果。

## 许可证

本项目以 [GNU AGPL-3.0](LICENSE) 发布，并采用 PyMuPDF 的 AGPL 许可选项。Microsoft Office、LibreOffice 及其他第三方组件仍遵循各自许可条款，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。发布便携版时，对应源码应保留在同一 GitHub 仓库和版本标签中。
