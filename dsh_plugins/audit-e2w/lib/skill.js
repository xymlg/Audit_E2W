export const skillName = "audit-word-excel"
export const skillDescription = "通过 DSH Agent 在本机完成 Word 与 Excel 审计办公任务：创建 Excel、同步回 Word、标注、占位符、模板批量报告、空白表删除恢复、差异检查和备份回滚。"
export const skillWhenToUse = "用户要求处理本地 Word/Excel 审计文件、导出表格、同步数据、生成报告或检查文档差异时使用。"
export const skillContent = `# Audit_E2W 本地审计工作流

先调用 \`audit_capabilities\` 查看操作、必填参数和示例，再调用 \`audit_engine_status\` 确认引擎就绪。所有文件路径必须是 Windows 绝对路径。

## 安全顺序

1. 用 \`audit_query\` 扫描、比对或列出候选，确认文件和影响范围。
2. 对支持预演的操作先调用 \`audit_run\` 且 \`dryRun: true\`。
3. 向用户说明将修改或生成的文件；得到明确确认后，再以 \`dryRun: false\` 调用。写操作会触发 DSH 审批，审批未通过不得绕过。
4. 读取返回的日志、结果和文件路径，向用户核验实际结果；不要仅凭工具调用成功就宣称业务结果正确。

## 常见映射

- Word 导出 Excel：\`export\`
- Excel 同步回 Word：\`writeback\`
- Word 与 Excel 差异：\`diff\`
- 标注导出/同步：\`marks\` / \`marks-sync\`
- 占位符工作表补齐/同步：\`placeholders\` / \`placeholders-sync\`
- 单模板报告：\`template-batch\`
- 多报告同步：\`batch-writeback\`
- 空白表扫描/删除：\`blank-scan\` / \`blank-delete\`
- 备份恢复：\`backup-list\` / \`backup-restore\`

不得覆盖模板或源 Excel；不得擅自设置 \`force: true\` 或关闭备份。遇到版本不一致、文件占用、警告或部分失败，应停止并报告，不要静默降级。
`
