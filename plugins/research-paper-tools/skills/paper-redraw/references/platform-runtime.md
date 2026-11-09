# 运行与可移植性

仅在运行环境、Windows 交付、路径或字体问题涉及本次任务时读取；不扫描所有语言、GPU、内存或求解器。

- 复用宿主提供或项目已选的实际解释器，检查本图需要的包能导入和使用即可。路径存在或包元数据存在不证明入口可运行；不要无依据安装另一整套环境。
- PowerShell 使用 `&` 调用带空格的可执行路径，参数分开传递；不要照搬 Bash 的 `export`、反斜线续行和 `chmod`。Python 安装包与运行使用同一解释器。
- 上层目录允许中文与空格。文件和函数用简短英文名称，避免保留名；脚本从自身位置解析必要输入，例如 Python 的 `Path(__file__).resolve().parent`，不依赖启动时的当前目录。
- Node ESM 使用 `fileURLToPath`；依赖须能从实际构建目录解析，`NODE_PATH` 不能代替 ESM 包解析。宿主包目录可在私有构建目录链接或使用 Windows junction，不能把整套依赖复制进交付目录。
- 用参数数组调用子进程，不拼接不可信 shell 字符串。文本按 UTF-8 读写；字体选择以实际导出后能显示中文、希腊字母、上下标和负号为准。
- 渲染工具使用实际发现的路径；没有某个首选渲染器时，尝试宿主已有的预览或应用，不直接将整条重建路线判定失败。
- 缩略预览不是高保真检查：Quick Look 等可能省略上下标、原生图表或超出预览范围的列；保留正确原生对象，通过结构检查和可信渲染确认。渲染器支持范围不明时，先验证一个涉及该属性的局部样例，不把一种软件的结果写成另一种软件的实测。
- 不修改全局 PATH、执行策略或系统设置来完成普通重绘。缓存、host helper、临时包、许可证检测等只按实际需要处理。

数据图的独立性应在不含原图的私有目录验证，含空格或中文路径有实际风险时补测该入口。只报告真实测试环境；Linux／macOS 运行不能写成 Windows 或 Microsoft Office 真机实测。

## Windows 原生环境

先区分 Windows 原生 PowerShell 与 WSL。两者分别使用各自的解释器、依赖和路径；不要直接混用 `C:\...` 与 `/mnt/c/...`，也不把 WSL 的成功运行写成 Windows 原生实测。

按本图路线准备依赖，优先复用宿主提供的环境：

| 使用部分 | 需要的环境 |
|---|---|
| Office 结构检查 | Python 3.10+，标准库即可 |
| 图像对照 | Python 3.10+、Pillow 9.1+ |
| 数据重绘／表格构建 | 按实际生成代码选择库；常见为 NumPy、Matplotlib，或支持 rich text 的 openpyxl 3.1.x |
| 自带 PPT 模板 | Node.js 20+、PptxGenJS 4.0.x，包能从构建目录解析 |
| 自带 Office 渲染器 | LibreOffice 和 Poppler 的 `pdftoppm.exe`；已有可用宿主渲染器时无需另装这一套 |

渲染器依次探测 `soffice.com/.exe` 和 ProgramW6432／ProgramFiles／ProgramFiles(x86) 中的常见安装位置，优先 console 入口；自定义安装目录用 `--soffice` 明确传入。Poppler 不假定在固定目录，用 `--pdftoppm` 或已确认的 PATH。仅安装 Microsoft Office 不会自动满足这个 LibreOffice 辅助脚本的依赖。

PowerShell 示例中的路径必须替换成实际发现并验证过的路径，不依赖 `python` 商店别名：

```powershell
$pythonExe = 'C:\实际路径\python.exe'
$skillDir = 'C:\实际技能目录\paper-redraw'
$workDir = 'D:\项目目录\work\Fig01'
$sofficeExe = 'C:\Program Files\LibreOffice\program\soffice.com'
$pdftoppmExe = 'C:\实际路径\pdftoppm.exe'
& $pythonExe -X utf8 (Join-Path $skillDir 'scripts\render_office.py') (Join-Path $workDir 'figure.pptx') --output-dir (Join-Path $workDir 'qa') --round 1 --soffice $sofficeExe --pdftoppm $pdftoppmExe
```

Node 包目录在项目外时，Windows 可用本地目录 junction，避免把创建符号链接的权限作为默认前提。只为当前私有构建目录创建，不覆盖已有 `node_modules`，不修改全局环境：

```powershell
$moduleLink = Join-Path $buildDir 'node_modules'
# $buildDir 和 $runtimeNodeModules 均为实际发现的绝对目录。
if (-not (Test-Path -LiteralPath $moduleLink)) {
    New-Item -ItemType Junction -Path $moduleLink -Value $runtimeNodeModules | Out-Null
}
```

Windows 字体按本机实际安装情况选择，中文可检查 Microsoft YaHei／SimSun；不依赖 macOS 的 PingFang。检查希腊字母、上下标和负号的真实显示。中文／空格目录可正常使用；若第三方工具确有长路径限制，改用较短的私有构建路径，不改系统注册表或用户原件。

## Windows 验证范围

可将 `PAPER_REDRAW_NODE`、`PAPER_REDRAW_NODE_MODULES`、`PAPER_REDRAW_SOFFICE`、`PAPER_REDRAW_PDFTOPPM` 设置为该机实际路径，再在技能目录执行 `python -X utf8 -B -m unittest discover -s tests -v`。测试会使用含中文与空格的目录；Windows 上的 Node 测试使用 junction，结束时只移除链接。缺依赖而跳过的项目必须单独报告。

本技能已有 Windows 路径分支、中文编码和参数构造检查；在非 Windows 主机执行这些检查，只能证明相应分支逻辑。正式称为“Windows 已实测”前，仍须在 Windows 原生环境执行真实 PPT／Excel 构建与渲染，并查看产物。
