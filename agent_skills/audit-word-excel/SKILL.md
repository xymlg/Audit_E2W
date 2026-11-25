---
name: audit-word-excel
display_name: Audit_E2W
display_name_en: Audit Toolbox
description: 当用户要求用本地 Word 和 Excel 完成审计办公任务时使用，包括创建 Excel、同步回 Word、链接标注、正文占位符、模板生成报告、批量同步、差异检查、表格维护和备份恢复。遵循先预演或差异检查、写盘前确认、执行后核验的流程。
description_zh: 在本地处理 Word 与 Excel 的创建、同步、标注、占位符、批量报告、差异检查、表格维护和备份恢复，并由 Agent 先预演、确认再执行。
description_en: Create and synchronize local Word and Excel audit workpapers, marks, placeholders, batch reports, table maintenance, and backups with preview, approval, and result verification.
version: 1.11.1
author: xymlg
allowed-tools: Bash
examples_zh:
- "把这份 Word 报告导出成 Excel，先预演。"
- "Excel 改好了，只把表格同步回 Word，先检查差异。"
- "用这份模板和底稿生成一份新报告，不要覆盖原件。"
examples_en:
- "Export this Word report to Excel in preview mode."
- "Check the differences, then sync only the tables back to Word."
- "Generate a new report from this template and workbook without overwriting the originals."
---

# Word/Excel 审计任务

## 目标

把自然语言需求转换成正确的审计工具箱操作，安全完成本地文件处理，并用白话汇报结果。业务计算和文件修改由本地审计引擎完成，不凭模型猜测处理结果。

## 选择执行通道

### Qoder 插件通道

如果 `audit_engine_status`、`audit_query`、`audit_run`、`audit_task_status` 可用：

1. 调用 `audit_engine_status`，确认引擎可用且版本一致。
2. 调用 `audit_capabilities` 核对 operation、必填参数、默认值、预演支持和写盘标志；工具不可用时读取 `references/operations.md`。
3. 只读操作用 `audit_query`，执行操作用 `audit_run`。
4. 按返回的 `taskId` 用 `audit_task_status` 等待完成，不重复启动同一任务。

### 本地脚本通道

如果 Qoder 工具不可用，但当前平台提供 Bash、本地文件访问权限和已安装的审计工具箱引擎，按 `references/local-cli.md` 使用本 Skill 的 `scripts/audit_tool.py`。

先运行 `status`，只有引擎可用且 `versionMatched=true` 才继续。脚本一次调用对应一个请求，返回引擎日志和最终 `data`。长任务使用平台提供的后台执行能力，命令结束前不得重复执行。

如果两条通道都不可用，停止并说明缺少哪个能力。不要声称已经处理文件，也不要用普通文本编辑模拟 Word/Excel 同步。

## 输入与安全边界

- 只处理用户明确选择的文件，或用户明确授权的目录内文件；不要猜测客户文件名。
- 路径必须使用当前系统的绝对路径，扩展名只支持 `.docx`、`.docm`、`.xlsx`、`.xlsm`。
- 写盘前提醒用户关闭对应 Word/Excel 文件；默认 `force=false`。
- 查询操作只读。执行操作可能修改或生成文件，必须先预演或做相关只读检查。
- 删除、覆盖、恢复备份属于高风险动作，必须展示明确范围并取得确认。
- 不上传报告内容，不在公开消息、日志或示例中写客户名称、账号、密码或 API Key。

## 意图映射

| 用户说法 | operation | 关键参数或默认值 |
| --- | --- | --- |
| Word 导出 Excel、创建底稿 | `export` | `docx`、`excel`；三个工作表始终创建；默认写定位链接和合计公式 |
| Excel 同步回 Word、只改表格 | `writeback` | `docx`、`excel`；`contents.tables=true`，其他默认关闭 |
| 全部修改都同步 | `writeback` | `contents.tables=true`、`marks=true`、`placeholders=true` |
| 只看 Excel 和 Word 有什么差异 | `diff` | `docx`、`excel`；只读 |
| 导出差异清单 | `diff-export` | `docx`、`excel`；不修改 Word 正文 |
| 导出或刷新链接标注 | `marks` | `docx`、`excel` |
| 同步标注到 Word | `marks-sync` | 默认 `keepHighlight=false`；用户要求保留外观时传 `true` |
| 补齐正文占位符工作表 | `placeholders` | `docx`、`excel` |
| 同步正文占位符 | `placeholders-sync` | `docx`、`excel` |
| 检查模板有哪些字段 | `template-scan` | `docx`；只读 |
| 单模板生成、用模板和 Excel 出新报告 | `template-batch` | `templates`、`items=[{excel,template:1}]`、`outDir`；生成新 Word |
| 批量同步多份报告 | `batch-writeback` | `paths` 为 Word 清单；默认只同步表格 |
| 扫描空白表 | `blank-scan` | `docx`；只读 |
| 删除空白表 | `blank-delete` | 先扫描，再用候选编号 `pick`；默认可逆 |
| 恢复已删表格 | `restore-tables` | 先 `deleted-list`，再传恢复编号 `pick` |
| 查看或设置排序 | `sort-list`；`sort-set` | `code`、`column`、`descending` |
| 重建或清理书签 | `bookmark`、`clear-bookmarks` | 高风险维护；执行前说明影响 |
| 查看或恢复备份 | `backup-list`；`backup-restore` | 恢复会覆盖目标，必须明确备份文件和目标文件 |

完整参数见 `references/operations.md`。不要凭记忆猜参数。

## 消解歧义

- “同步到 Word”默认指 `writeback`，不是模板生成。
- “单模板”指 `template-batch`：用模板副本生成新 Word，不覆盖模板原件和源 Excel。
- “只同步表格”只设 `contents.tables=true`；“全部同步”才同时开启表格、标注、占位符。
- 用户说“生成报告”但没说是覆盖还是新建时，先确认，不要自行覆盖原报告。
- 用户说“这个文件夹”但没给文件规则时，先列出候选文件并确认，不要全部自动处理。

## 标准工作流程

### 1. 只读预检

- 导出前确认 Word 存在，并检查表格、标注或占位符的扫描结果。
- 同步前调用 `diff` 或支持的预演，先得到改动、警告和可能失败项。
- 模板生成前调用 `template-scan` 检查占位符；存在多份 Excel 时逐份核对模板归属。
- 删表或恢复前先扫描并展示候选编号，不凭印象选择。

### 2. 先预演

以下操作支持 `dryRun=true`：

`export`、`writeback`、`bookmark`、`rejudge`、`clear-bookmarks`、`marks`、`marks-sync`、`placeholders`、`placeholders-sync`、`blank-delete`、`restore-tables`、`blank-batch-delete`、`batch-writeback`、`template-batch`。

不支持的写盘操作，用相关只读查询代替预检，并在执行前明确说明。不要把 `dryRun=true` 当成所有操作都有效。

### 3. 向用户确认

用三到五行白话说明：

- 准备处理哪些文件；
- 会修改或生成哪些文件；
- 将同步哪些内容；
- 有哪些警告、跳过项或备份安排。

只有用户明确说“确认”“直接执行”“不用再问”，或已对当前范围给出明确执行指令时，才继续写盘。用户只要求“先看看”时，不得写盘。

### 4. 正式执行

- 把预演请求中的 `dryRun` 改为 `false`。
- 默认开启 `backup=true`；只有用户明确要求时才关闭备份。
- 默认 `force=false`；只有用户明确同意跳过占用检查时才传 `true`。
- `.xlsm` 必须保留 VBA。
- 模板生成必须输出到 `outDir` 的新文件，不覆盖模板和 Excel。
- 批量任务逐份报告成功、失败和跳过情况。

### 5. 等待并判断完成

Qoder 通道使用 `audit_task_status` 持续查询，直到 `succeeded`、`failed` 或 `cancelled`。`succeeded` 只表示引擎执行成功；批量任务仍须检查 `summary.outcome` 和 `summary.batch`，可能是部分失败。

本地脚本通道以命令退出码和返回 JSON 的 `ok` 为准。`ok=false` 时读取 `error`、`logs` 和 `stderr`，不要只看见进程结束就报成功。

### 6. 核验并汇报

至少核对：

- 输出文件是否存在；
- 模板和源 Excel 是否保持未修改；
- 警告、跳过项、失败清单和需要复核的日志；
- 改动数量、标注数、占位符数是否与预演一致；
- 备份路径是否明确。

Python operation 返回整数 `0` 常表示成功码，不是“零个改动”。数量以引擎日志和任务汇总为准。Qoder 的 `summary.needsReview=true` 或 `outcome=partial` 时，必须逐项说明，不能只报“已完成”。

最终按以下格式汇报：

> 已完成：  
> 生成/修改文件：  
> 主要改动：  
> 警告或跳过项：  
> 备份位置：  
> 建议复核：

## 失败处理

- 文件被占用：请用户关闭文件后重试，不自动强制保存。
- 路径或文件不存在：指出准确路径，请用户补充或更正。
- 书签无法匹配：提示工作簿可能不是从该模板导出，先检查来源或重建书签。
- 占位符没有填写：列出缺少的字段，不编造值。
- 公式没有缓存结果：提醒先在 Excel 中重算并保存。
- 执行失败：读取具体日志；不要原样反复重试。
- 需要回滚：先用 `backup-list` 查看备份，再确认恢复目标。
- 批量结果中有失败或跳过时，逐份列出文件和原因。

## 平台资源

- 操作参数：`references/operations.md`
- 本地脚本配置：`references/local-cli.md`
- Qoder、WorkBuddy、DSH 安装位置：`references/platforms.md`
