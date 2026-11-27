---
description: 按样板或格式方案调整报告格式，候选确认后先生成新副本
---
使用 `audit-word-excel` Skill 完成“报告格式调整”任务。

本次目标：按用户指定的样板或格式方案调整报告格式，并生成新的 Word 副本。不得覆盖原报告。

1. 调用 `audit_engine_status`，确认引擎可用且版本一致；调用 `audit_capabilities` 核对 `format-scan`、`format-heading-candidates` 和 `format-apply`。
2. 需要识别标题、表题或附注边界时，只用 `audit_query` 返回的候选和有限上下文判断；展示建议、置信度和待确认项，不得读取或上传整篇报告来补充。
3. 用户确认候选映射后，用原批次编号调用 `format-scan` 预演；展示七项调整范围和 `visual_review` 清单。
4. 预演后写盘前等待用户明确确认，再用 `audit_run` 调用 `format-apply`，指定新的输出路径。
5. 用 `audit_task_status` 等待结束，核对输出为新文件、原报告未修改、确认数量一致，并提醒用户在 Word 中目视复核。

斜杠指令不是写盘确认；不得自动设置 `confirmed=true`、`force=true` 或关闭备份。正式应用前必须重新取得用户对预演结果的确认。不要用 Bash 直接调用引擎。
