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
| `append-table` | 追加新表块 | `docx`、`name`、`nrows`、`ncols` | `excel`、`after` | 否 |
| `sort-set` | 设置表格排序 | `docx`、`code`、`column` | `excel`、`descending` | 否 |
| `sort-clear` | 清除表格排序 | `docx`、`code` | `excel` | 否 |
| `batch-writeback` | 多报告批量同步 | `paths` | `contents`、`backup`、`keepHighlight`、`force` | 是 |
| `template-batch` | 模板批量生成 | `templates`、`items`、`outDir` | `contents`、`keepHighlight` | 是 |
| `fill-placeholders` | 生成正文信息填写表 | `docx`、`paths` | 无 | 否 |
| `backup-restore` | 恢复历史备份 | `backupPath`、`target` | 无 | 否 |

## 参数约定

- `dryRun=true`：只预演，不写文件；正式执行时改为 `false`。
- `backup=true`：覆盖前备份到原文件旁 `_backup`，默认开启。旧参数 `noBackup=true` 表示不备份，不要与新参数混用。
- `force=true`：跳过文件占用检查，仅在用户明确同意时使用。
- `contents`：可包含 `tables`、`marks`、`placeholders`、`keepHighlight` 布尔值；三项同步内容至少选择一项。
- `pick`：候选或恢复编号，使用扫描结果中的原始文本，不自行重编号。
- `nrows`：1～200；`ncols`：1～30。
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
