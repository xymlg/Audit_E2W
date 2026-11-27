# 操作参数

所有路径使用绝对路径。`excel` 未传时通常使用 Word 同目录的 `<主名>_表格数据.xlsx`。批量清单最多 200 项。

## 只读查询

| operation | 用途 | 必填 | 可选 |
| --- | --- | --- | --- |
| `blank-scan` | 扫描空白表 | `docx` | `docx` |
| `deleted-list` | 查看可恢复表格 | `docx` | `docx`、`excel` |
| `blank-batch-scan` | 批量扫描空白表 | `paths` | `paths` |
| `diff` | Excel 与 Word 差异 | `docx` | `docx`、`excel` |
| `table-list` | 查看工作簿表块 | `docx` | `docx`、`excel` |
| `sort-list` | 查看排序设置 | `docx` | `docx`、`excel` |
| `template-scan` | 检查模板占位符 | `docx` | `docx` |
| `backup-list` | 查看历史备份 | `docx` | `docx`、`excel` |
| `format-scan` | 格式调整预演 | `docx` | `template` 或 `profileName`、`scope`、`headingBatchId`、`headingLevels`、`roleBatchId`、`roleOverrides`、`regionStartId` |
| `format-heading-candidates` | 筛选候选标题、表格前段落角色及附注边界 | `docx` | `docx` |

查询操作不支持 `dryRun`，不要添加该参数。

## 执行操作

| operation | 用途 | 必填 | 常用可选参数 | 支持预演 |
| --- | --- | --- | --- | --- |
| `export` | 创建 Excel | `docx` | `excel`、`backup`、`withExcelLinks`、`totalsFormula`、`force` | 是 |
| `writeback` | 同步到 Word | `docx` | `excel`、`contents`、`backup`、`deleteEmptyRows`、`changelog`、`force` | 是 |
| `bookmark` | 重建或迁移书签 | `docx` | `excel`、`withExcelLinks`、`force` | 是 |
| `rejudge` | 重新判定 ✓/✕ | `docx` | `excel`、`backup`、`force` | 是 |
| `clear-bookmarks` | 清理工具书签 | `docx` | `force` | 是 |
| `marks` | 导出或刷新标注 | `docx` | `excel`、`force` | 是 |
| `marks-sync` | 同步标注到 Word | `docx` | `excel`、`keepHighlight`、`force` | 是 |
| `placeholders` | 补齐占位符工作表 | `docx` | `excel`、`force` | 是 |
| `placeholders-sync` | 同步占位符到 Word | `docx` | `excel`、`force` | 是 |
| `blank-delete` | 删除空白表格 | `docx`、`pick` | `excel`、`keepCaption`、`keepCaptionFor`、`reversible`、`force` | 是 |
| `restore-tables` | 恢复已删表格 | `docx`、`pick` | `excel`、`force` | 是 |
| `blank-batch-delete` | 批量删除空白表 | `items` | `force` | 是 |
| `diff-export` | 导出差异清单 | `docx` | `excel` | 否 |
| `incremental-export` | 增量追加表格/标注/占位符 | `docx` | `excel`、`backup`、`withExcelLinks`、`totalsFormula`、`force` | 是 |
| `sort-set` | 设置表格排序 | `docx`、`code`、`column` | `excel`、`descending` | 否 |
| `sort-clear` | 清除表格排序 | `docx`、`code` | `excel` | 否 |
| `batch-writeback` | 多报告批量同步 | `paths` | `contents`、`backup`、`keepHighlight`、`force` | 是 |
| `template-batch` | 模板批量生成 | `templates`、`items`、`outDir` | `contents`、`keepHighlight` | 是 |
| `fill-placeholders` | 生成正文信息填写表 | `docx`、`paths` | 无 | 否 |
| `backup-restore` | 恢复历史备份 | `backupPath`、`target` | 无 | 否 |
| `format-apply` | 应用报告格式 | `docx` | `template` 或 `profileName`、`output`、`scope`、`headingBatchId`、`headingLevels`、`roleBatchId`、`roleOverrides`、`regionStartId`、`confirmed` | 是 |

## 写盘确认

所有 `audit_run` 写操作在 `dryRun=false` 时都必须同时传 `confirmed=true`，表示用户已确认预演结果和受影响文件。服务端和引擎都会拒绝缺少确认的正式写请求；`dryRun=true` 不需要该参数。`format-apply` 还必须遵守下方的标题、角色和边界确认规则。

- `template` 是样板 Word 绝对路径；`profileName` 是已存纯格式方案名称，二者只能选一个。都未提供时按原报告自身格式生成副本。
- `scope` 提供时必须完整包含七个布尔键：`cover`、`headings`、`statement`、`body`、`table_text`、`table_emphasis`、`table_borders`；至少一项为 `true`。
- `output` 必须是新的 `.docx`/`.docm` 绝对路径，原报告和已有文件不会被覆盖。
- `format-apply` 的 `dryRun=true` 只预演；正式执行必须同时传 `dryRun=false`、`confirmed=true`。
- `format-heading-candidates` 最多返回 300 个标题/角色候选，另含最多 8 个附注边界候选；每个候选只含原文（最多 120 字）、前后各最多 100 字及有限格式线索，不含整篇报告、表格单元格、页眉页脚或批注。`truncated=true` 表示还有候选未返回。
- `headingBatchId` 与 `headingLevels` 必须同时提供。`headingLevels` 格式为 `{"候选id": 1}`，最多 300 项，级别为 1～4 整数；只放用户确认的候选，且 `scope.headings=true`。
- `roleBatchId` 与 `roleOverrides` 同时提供；角色映射只接受候选 `options` 中的 `body` 或 `table_title`，且 `scope.statement=true`。`regionStartId` 必须来自 `boundary_candidates` 并与 `roleBatchId` 同时提供。标题和角色同时应用时，两个批次编号必须相同。
- 报告或候选变化后旧批次失效，必须重新扫描并重新确认。标题和角色只改直接排版，边界只改变本次处理范围；不写标题样式、不改大纲级别、不更新目录。正式结果中的 `visual_review` 是 Word 目视验收清单，结构核验不等于视觉通过。

## 参数约定

- `dryRun=true`：只预演，不写文件；正式执行时改为 `false`。
- `backup=true`：覆盖前备份到原文件旁 `_backup`，默认开启。旧参数 `noBackup=true` 表示不备份，不要与新参数混用。
- `force=true`：跳过文件占用检查，仅在用户明确同意时使用。
- `contents`：可包含 `tables`、`marks`、`placeholders`、`keepHighlight` 布尔值；三项同步内容至少选择一项。
- `pick`：候选或恢复编号，使用扫描结果中的原始文本，不自行重编号。
- `incremental-export` 只追加尚未登记内容：新表进入「新增表格」XT01…，新标注 XM01…，新占位符追加「正文信息」；既有行不动。表格内高亮和占位符排除。无新增不写盘。
- `descending`：默认 `true` 表示降序，`false` 表示升序。
- `backup-restore` 的备份文件与目标文件扩展名必须一致。

## 批量清单

`template-batch.items` 中每项：

```json
{"excel": "<工作簿绝对路径>", "template": 1}
```

`template` 是从 1 开始的模板编号，对应 `templates` 数组位置。

`blank-batch-delete.items` 中每项：

```json
{
  "docx": "<Word 绝对路径>",
  "pick": "1",
  "keepCaptionFor": "2",
  "reversible": true,
  "excel": "<可选工作簿绝对路径>"
}
```

批量结果逐项检查 `status`、`reason` 或 `error`。顶层成功不代表每份文件都成功。
