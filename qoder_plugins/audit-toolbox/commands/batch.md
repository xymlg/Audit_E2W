---
description: 批量同步多份报告或用模板批量生成新报告，逐份核对结果
---
使用 `audit-word-excel` Skill 完成“批量同步”或“模板批量生成”任务。

本次目标：处理用户明确授权的文件清单或目录规则。先区分 `batch-writeback` 与 `template-batch`；目录规则不明确时先列出候选文件并确认，不得自动处理全部文件。

1. 调用 `audit_engine_status`，确认引擎可用且版本一致。
2. 调用 `audit_capabilities` 核对所选批量操作的必填参数、模板归属、输出目录和预演支持。
3. 先用 `audit_query` 完成只读检查，再按操作支持情况用 `audit_run` 发送 `dryRun=true`；展示完整文件清单、输出位置、跳过项和警告，写盘前等待用户明确确认。
4. 确认后用 `audit_run` 正式执行；模板生成必须输出到新目录或新文件，不覆盖模板和源 Excel。
5. 用 `audit_task_status` 等待结束，检查 `summary.outcome`、`summary.batch` 和 `needsReview`，逐份汇报成功、失败和跳过情况。

斜杠指令不是写盘确认；不得自动设置 `confirmed=true`、`force=true` 或关闭备份。不要用 Bash 直接调用引擎，不上传整篇报告。
