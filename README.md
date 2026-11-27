# Audit_E2W

让 Agent 按你的指令处理本机 Word / Excel 审计文档：从 Word 提取表格、标注和正文占位符，在 Excel 中修改，再同步回 Word；也支持空白表格处理、差异检查、排序、批量报告和备份恢复。

当前版本：**1.11.3**。仅支持 **Windows 10/11 64 位**，不支持 macOS。

## 能做什么

- **Word → Excel**：导出表格；只要表格、链接标注、正文占位符任意一种存在就导出，并创建三个工作表。
- **Excel → Word**：同步单元格、行列结构、标注、正文占位符及数值显示格式；同步前可检查差异。
- **表格维护**：重建书签、重新判定表块、增量导出、设置或清除排序。
- **空白表格**：扫描、删除、查看已删除表块、恢复；支持批量处理。
- **批量办公**：批量同步、单模板批量生成报告、批量填写占位符。
- **安全回退**：修改前备份，可查看备份历史并恢复；预演不写文件，危险操作需确认。

## 三种使用方式

| 方式 | 适用场景 | 需要什么 |
| --- | --- | --- |
| Windows 图形界面 | 直接手工操作 | Python 源码环境，或已构建的 Windows 程序 |
| Qoder 插件 | 在 Qoder 中用自然语言指挥 Agent，并提供操作界面 | Qoder Desktop、Windows 版引擎 |
| DSH 完整插件 | 在 DSH 中用自然语言指挥 Agent | DSH、Node.js 24.11+、pnpm 10+、Python 3.12+ |
| WorkBuddy / DSH Skill | 只安装工作流说明，由 Agent 调用本机执行脚本 | 对应 Agent、Windows 版引擎 |

### 1. 源码运行

在项目根目录安装依赖并启动图形界面：

```powershell
uv sync
.\.venv\Scripts\pythonw.exe .\审计工具箱.pyw
```

没有 uv 时，也可使用 Python 3.12+：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\pythonw.exe .\审计工具箱.pyw
```

命令行示例（在 `excel_word` 目录执行）：

```powershell
..\.venv\Scripts\python.exe .\word_excel_sync.py export .\报告.docx --dry-run
..\.venv\Scripts\python.exe .\word_excel_sync.py export .\报告.docx
..\.venv\Scripts\python.exe .\word_excel_sync.py writeback .\报告.docx --dry-run
..\.venv\Scripts\python.exe .\word_excel_sync.py writeback .\报告.docx
```

### 2. Qoder 插件

源码目录为 `qoder_plugins/audit-toolbox`。构建脚本会把 Windows 引擎、MCP Server、界面和 Skill 打入插件 ZIP：

```powershell
.\.venv\Scripts\python.exe .\packaging\build_qoder_plugin.py --check-only
.\.venv\Scripts\python.exe .\packaging\build_qoder_plugin.py
```

构建前需要 PyInstaller；产物位于 `dist/qoder_plugins`。Qoder 插件包使用 Qoder 内置 Node.js 运行 MCP Server。构建脚本按 `packaging/build_qoder_plugin.py` 的明确清单复制插件源文件，拒绝链接/junction、客户 Office 文档、锁文件和缓存，并在生成 ZIP 后复核同一清单；`--check-only` 同时核对版本、源树、安装脚本白名单及已有便携版目录。

### 3. DSH 完整插件

仓库根目录是可安装的 DSH npm bundle，入口声明在 `cordis.patch.yml`，提供四个工具：

- `audit_capabilities`：查看 31 个操作及参数。
- `audit_engine_status`：检查本机引擎和版本。
- `audit_query`：只读扫描与差异检查。
- `audit_run`：执行操作；非预演写操作必须经 DSH 用户审批。

安装包优先使用 `AUDIT_TOOLBOX_ENGINE` 指定的 Windows 引擎，其次使用包内 `engine/审计工具箱引擎.exe`，最后使用 Python 源码引擎。公开 npm/GitHub 包**不附带 Windows exe**，因此源码方式需要先安装 Python 依赖：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install python-docx==1.2.0 openpyxl==3.1.5 lxml==6.1.1
```

可通过 `AUDIT_TOOLBOX_PYTHON` 指定解释器的绝对路径。插件与引擎版本不一致时会拒绝执行，不会静默降级。

安装到 DSH profile（Windows）：

```powershell
npm install -g pnpm@10
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install python-docx==1.2.0 openpyxl==3.1.5 lxml==6.1.1
$env:AUDIT_TOOLBOX_PYTHON = (Resolve-Path .\.venv\Scripts\python.exe).Path

dsh plugin --profile desktop add github:xymlg/Audit_E2W
dsh --profile desktop --dump-config
dsh --profile desktop
```

`desktop` 可替换为实际 profile 名；`--dump-config` 应显示 `audit-e2w` 配置层。GitHub 安装的是纯 JavaScript 发布入口，不需要构建脚本授权。生产环境建议把安装地址锁定到具体提交，例如 `github:xymlg/Audit_E2W#<commit-sha>`。

开发和打包检查：

```powershell
npm run check
npm test
npm pack --dry-run
```

### 4. 通用 Agent Skill

`agent_skills/audit-word-excel-workbuddy` 是 WorkBuddy 专用工作流 Skill，只走本地脚本通道；`agent_skills/audit-word-excel` 是 DSH 等 Agent 的通用工作流 Skill。两者都不包含 Windows 引擎二进制，必须配合本机执行脚本或已安装的 Windows 程序使用。

```powershell
.\.venv\Scripts\python.exe .\packaging\build_agent_skills.py --check-only
.\.venv\Scripts\python.exe .\packaging\build_agent_skills.py --build-zips
```

## 推荐工作流程

1. 让 Agent 先检查引擎状态和文件路径。
2. 先预演、扫描或比对差异，不直接写盘。
3. 说明将修改或生成哪些文件。
4. 获得确认后执行；写操作通过宿主审批或由用户明确授权。
5. 核对结果、警告和部分失败项；必要时从备份恢复。

示例：

> 先检查引擎状态，然后对 `D:\工作\报告.docx` 执行导出预演，告诉我会生成哪些工作表和文件，不要写盘。

> 对 `D:\工作\报告.docx` 和 `D:\工作\报告_表格数据.xlsx` 做差异检查，列出变化，等我确认后再同步。

## 安全与数据边界

- 所有文件操作在本机执行，不上传报告内容。
- 插件不注册 HTTP 服务，不提供远程文件访问接口。
- Agent 工具要求 Windows 绝对路径，并限制文件扩展名。
- 写操作、取消、版本校验和 JSONL 引擎通信均有自动化测试。
- 请优先使用副本；修改前备份不能替代独立归档。
- 本项目不包含客户报告、账号、密钥或二进制构建产物。

## 开发

```powershell
uv sync
.\.venv\Scripts\python.exe -m pytest excel_word/tests -q
npm --prefix qoder_plugins/audit-toolbox run check
npm --prefix qoder_plugins/audit-toolbox run smoke
npm test
```

Python 测试会在项目内 `.tmp` 下创建并自动清理测试文档。目录结构和扩展约定见 `功能规划与文件布局.md`；详细接口见 `excel_word/README.md`。

## 许可证

[MIT](LICENSE)。
