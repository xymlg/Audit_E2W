export const skillName = "audit-word-excel"
export const skillDescription = "通过 DSH Agent 在本机完成 Word 与 Excel 审计办公任务：创建 Excel、同步回 Word、标注、占位符、模板批量报告、空白表维护、报告格式调整、差异检查和备份回滚。"
export const skillWhenToUse = "用户要求处理本地 Word/Excel 审计文件、导出表格、同步数据、生成报告、调整报告格式或检查文档差异时使用。"
export const skillContent = `# Audit_E2W 本地审计工作流

先调用 \`audit_capabilities\` 查看全部 31 个操作、必填参数和示例，再调用 \`audit_engine_status\` 确认引擎就绪且版本一致。所有文件路径必须是 Windows 绝对路径。

## 安全顺序

1. 用 \`audit_query\` 扫描、比对或列出候选，确认文件和影响范围。
2. 对支持预演的操作先调用 \`audit_run\` 且 \`dryRun: true\`，展示目标文件、改动数量、警告和跳过项。
3. 向用户说明将修改或生成的文件；得到明确确认后，再以 \`dryRun: false\`、\`confirmed: true\` 调用。写操作会触发 DSH 审批，审批未通过不得绕过。
4. 读取返回的日志、结果和文件路径，向用户核验实际结果；不要仅凭工具调用成功就宣称业务结果正确。
5. 批量任务逐份汇报成功、失败和跳过；顶层成功不代表每份文件都成功。

## 常见映射

- Word 导出 Excel：\`export\`
- Excel 同步回 Word：\`writeback\`
- Word 与 Excel 差异：\`diff\`
- 标注导出/同步：\`marks\` / \`marks-sync\`
- 占位符工作表补齐/同步：\`placeholders\` / \`placeholders-sync\`
- 单模板报告：\`template-batch\`
- 多报告同步：\`batch-writeback\`
- 空白表扫描/删除：\`blank-scan\` / \`blank-delete\`
- 备份查看/恢复：\`backup-list\` / \`backup-restore\`
- 格式调整预演：\`format-scan\`
- 候选标题、表题角色与附注边界：\`format-heading-candidates\`
- 应用报告格式：\`format-apply\`

## 报告格式

1. 用户要用样板调整格式时，把 \`docx\` 和样板 \`template\` 传给 \`format-scan\`；\`template\` 与已存 \`profileName\` 二选一。未指定范围时默认七项全选；自定义 \`scope\` 必须完整包含 \`cover\`、\`headings\`、\`statement\`、\`body\`、\`table_text\`、\`table_emphasis\`、\`table_borders\`，不能全关。
2. 需要识别未标注的一至四级标题、区分表题与表格前引导正文、核对附注边界时，调用 \`format-heading-candidates\`。只使用返回的 \`candidates\`、\`role_candidates\`、\`boundary_candidates\` 和有限上下文，不得读取或上传整篇报告、表格单元格、页眉页脚或批注来补充判断。\`truncated=true\` 时说明候选已达上限，不能声称覆盖全部疑难项。
3. 根据候选原文和有限格式线索给出标题级别、角色和边界建议；标题级别只能是 1～4，角色只能是 \`body\` 或 \`table_title\`。逐项展示置信度，不确定的列为待确认，只在用户确认后保存 \`headingLevels\`、\`roleOverrides\`、\`regionStartId\`。
4. 将原样确认的 \`headingBatchId\`、\`headingLevels\` 和相同的 \`docx\`/样板/范围传给 \`format-scan\`。角色或边界使用同一个批次作为 \`roleBatchId\`；标题与角色同时确认时两个批次编号必须一致。核对预演数量、\`assigned_headings\`、\`assigned_roles\` 和 \`region_start\`；报告或候选变化导致旧批次被拒绝时，重新扫描并重新确认。
5. 预演后取得用户执行确认，调用 \`format-apply\`：\`dryRun=true\` 只预演；正式执行用 \`dryRun=false\`、\`confirmed=true\`，并指定新的 \`output\` 路径。绝不覆盖原报告。
6. 候选级别和角色只应用直接排版，边界只改变本次格式处理范围；不写标题样式、不改大纲级别、不更新目录。完成后核对输出为新文件、原报告未改、确认数量一致，并按返回的 \`visual_review\` 提醒用户在 Word 中复核封面、标题、正文、表题、宽表格和分页。结构核验通过不等于目视排版通过。

不得覆盖模板或源 Excel；不得擅自设置 \`force: true\` 或关闭备份。遇到版本不一致、文件占用、警告或部分失败，应停止并报告，不要静默降级。
`
