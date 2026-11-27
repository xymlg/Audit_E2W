---
description: 将指定 Excel 的改动同步回 Word，先检查差异再确认写回
---
使用 `audit-word-excel` Skill 完成“Excel 同步回 Word”任务。

本次目标：把用户指定的 Excel 中已修改内容同步到配对的 Word。先核对 Word 与 Excel 是否对应；同步范围不明确时先展示差异，不默认全选。

1. 调用 `audit_engine_status`，确认引擎可用且版本一致。
2. 调用 `audit_capabilities` 核对 `diff`、`writeback` 参数和 `contents` 默认值。
3. 先用 `audit_query` 执行差异检查；需要预演时用 `audit_run` 发送 `dryRun=true`，展示将同步的表格、标注、占位符及警告。
4. 向用户确认文件配对和同步范围，写盘前等待用户明确确认；确认后才用 `audit_run` 正式执行。
5. 用 `audit_task_status` 等待结束，核对改动数量、目标文件、警告和备份位置；失败或跳过项逐项说明。

斜杠指令不是写盘确认；不得自动设置 `confirmed=true`、`force=true` 或关闭备份。不要用 Bash 直接调用引擎，不上传整篇报告。
