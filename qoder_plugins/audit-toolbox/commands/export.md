---
description: 将选定的 Word 审计报告导出为 Excel，先预演再写出
---
使用 `audit-word-excel` Skill 完成“Word 导出 Excel”任务。

本次目标：把用户指定的 `.docx` / `.docm` 导出为新的 `.xlsx` / `.xlsm` 工作簿。文件缺失时请用户补充绝对路径，不得猜测客户文件。

1. 调用 `audit_engine_status`，确认引擎可用且版本一致。
2. 调用 `audit_capabilities` 核对 `export` 参数、默认值和预演支持；只读检查用 `audit_query`，预演或执行用 `audit_run`。
3. 先预演，展示输入文件、输出路径、预计生成的工作表和警告，写盘前等待用户明确确认。
4. 确认后正式执行，保留默认备份和文件占用检查；用 `audit_task_status` 等待结束，不重复启动任务。
5. 核对输出文件存在、三个工作表齐全、原报告未修改，再汇报警告、跳过项和备份位置。

斜杠指令不是写盘确认；不得自动设置 `confirmed=true`、`force=true` 或关闭备份。用户只要求“先看看”时仅返回预演。不要用 Bash 直接调用引擎，不上传整篇报告。
