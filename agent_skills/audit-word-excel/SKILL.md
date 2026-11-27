---
name: audit-word-excel
display_name: Audit_E2W
display_name_en: Audit Toolbox
description: 当用户要求用本地 Word 和 Excel 完成审计办公任务时使用，包括创建 Excel、同步回 Word、链接标注、正文占位符、模板生成报告、批量同步、差异检查、表格维护、报告格式调整和备份恢复。遵循先预演或差异检查、写盘前确认、执行后核验的流程。
description_zh: 在本地处理 Word 与 Excel 的创建、同步、标注、占位符、批量报告、差异检查、表格维护、报告格式调整和备份恢复，并由 Agent 先预演、确认再执行。
description_en: Create and synchronize local Word and Excel audit workpapers, marks, placeholders, batch reports, table maintenance, report formatting, and backups with preview, approval, and result verification.
version: 1.11.3
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
- 查询操作只读。执行操作可能修改或生成文件，必须先预演或做相关只读检查。所有正式写操作都必须同时传 `dryRun=false`、`confirmed=true`；服务端会拒绝缺少确认的请求。
- 删除、覆盖、恢复备份属于高风险动作，必须展示明确范围并取得确认。
- 不上传整篇报告，不在公开消息、日志或示例中写客户名称、账号、密码或 API Key。标题识别仅允许把 `format-heading-candidates` 返回的候选标题及有限上下文交给当前宿主 Agent；不得额外读取或发送全文、表格、页眉页脚、批注等内容。

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
| 按样板或格式方案预演 | `format-scan` | `docx`、`template` 或 `profileName`、七项 `scope`；只读 |
| 识别未标注标题、表题与附注边界 | `format-heading-candidates` | `docx`；只返回本地候选和有限上下文 |
| 应用格式或确认的标题级别 | `format-apply` | 先预演；输出新副本；正式执行传 `dryRun=false`、`confirmed=true` |

完整参数见 `references/operations.md`。不要凭记忆猜参数。

## 消解歧义

- “同步到 Word”默认指 `writeback`，不是模板生成。
- “单模板”指 `template-batch`：用模板副本生成新 Word，不覆盖模板原件和源 Excel。
- “只同步表格”只设 `contents.tables=true`；“全部同步”才同时开启表格、标注、占位符。
- 用户说“生成报告”但没说是覆盖还是新建时，先确认，不要自行覆盖原报告。
- 用户说“这个文件夹”但没给文件规则时，先列出候选文件并确认，不要全部自动处理。

## 报告格式与候选标题

1. 用户要用样板调整格式时，把 `docx` 和样板 `template` 传给 `format-scan`；`template` 与已存 `profileName` 二选一。未指定范围时默认七项全选；自定义 `scope` 必须完整包含 `cover`、`headings`、`statement`、`body`、`table_text`、`table_emphasis`、`table_borders`，不能全关。
2. 用户要求识别未标注的一至四级标题，或区分表题与表格前引导正文、核对附注边界时，调用 `format-heading-candidates`。只使用返回的 `candidates`（标题）、`role_candidates`（表格前段落）和 `boundary_candidates`（附注起始边界），不要读取全文来补充。向用户展示候选原文、编号和有限上下文；`truncated=true` 时说明候选已达 300 项上限，不能声称已覆盖全部疑难项。
3. 根据候选的 `text`、`before`、`after`、`style`、`bold`、`size`、`align`、`numbered` 判断标题级别；只使用 `role_candidates[].options` 判断 `body`/`table_title`；只从 `boundary_candidates` 选择附注起始段。编号、居中和加粗只是线索；普通正文、报表名称、日期和目录条目不得仅凭外观判为标题。给出高/中/低置信度，不确定的候选列为“待确认”，不要强行归类。
4. 输出建议映射：标题为 `{"候选id": 1}`，级别只能是 1～4 整数；疑难角色为 `{"候选id": "body"}` 或 `{"候选id": "table_title"}`；边界为单个 `regionStartId`。不是标题或不改角色的候选不列入。逐项展示建议、置信度和边界变化，允许修改或删除，只在用户确认后保存 `headingLevels`、`roleOverrides`、`regionStartId`。
5. 将原样确认的 `headingBatchId`、`headingLevels` 和相同的 `docx`/样板/范围传给 `format-scan`。角色或边界使用同一个 `batch_id` 作为 `roleBatchId`；标题与角色同时确认时两个批次编号必须一致。核对预演数量、`assigned_headings`、`assigned_roles` 和 `region_start`；报告或候选变化导致旧批次被拒绝时，重新扫描并重新确认，不得沿用旧结果。
6. 预演后取得用户执行确认，调用 `format-apply`：`dryRun=true` 只预演；正式执行用 `dryRun=false`、`confirmed=true`，并指定新的 `output` 路径。绝不覆盖原报告。
7. 候选级别和角色只应用直接排版，边界只改变本次格式处理范围；不写标题样式、不改大纲级别、不更新目录。完成后核对输出为新文件、原报告未改、确认数量一致，并按返回的 `visual_review` 提醒用户在 Word 中复核封面、标题、正文、表题、宽表格和分页。结构核验通过不等于目视排版通过。

## 标准工作流程

### 1. 只读预检

- 导出前确认 Word 存在，并检查表格、标注或占位符的扫描结果。
- 同步前调用 `diff` 或支持的预演，先得到改动、警告和可能失败项。
- 模板生成前调用 `template-scan` 检查占位符；存在多份 Excel 时逐份核对模板归属。
- 删表或恢复前先扫描并展示候选编号，不凭印象选择。

### 2. 先预演

以下操作支持 `dryRun=true`：

`export`、`writeback`、`bookmark`、`rejudge`、`clear-bookmarks`、`marks`、`marks-sync`、`placeholders`、`placeholders-sync`、`blank-delete`、`restore-tables`、`blank-batch-delete`、`batch-writeback`、`template-batch`、`format-apply`。

不支持的写盘操作，用相关只读查询代替预检，并在执行前明确说明。不要把 `dryRun=true` 当成所有操作都有效。

### 3. 向用户确认

用三到五行白话说明：

- 准备处理哪些文件；
- 会修改或生成哪些文件；
- 将同步哪些内容；
- 有哪些警告、跳过项或备份安排。

只有用户明确说“确认”“直接执行”“不用再问”，或已对当前范围给出明确执行指令时，才继续写盘。用户只要求“先看看”时，不得写盘。

### 4. 正式执行

- 把预演请求中的 `dryRun` 改为 `false`。`format-apply` 还必须传 `confirmed=true`。
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
>
> 生成/修改文件：
>
> 主要改动：
>
> 警告或跳过项：
>
> 备份位置：
>
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
- 平台安装位置与边界：`references/platforms.md`
