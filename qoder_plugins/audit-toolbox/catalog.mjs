import path from "node:path"

export const emptySchema = { type: "object", properties: {}, additionalProperties: false }

export const queryOperations = ["blank-scan", "deleted-list", "blank-batch-scan", "diff", "table-list", "sort-list", "template-scan", "backup-list"]
export const runOperations = ["export", "writeback", "bookmark", "rejudge", "clear-bookmarks", "marks", "marks-sync", "placeholders", "placeholders-sync", "blank-delete", "restore-tables", "blank-batch-delete", "diff-export", "append-table", "sort-set", "sort-clear", "batch-writeback", "template-batch", "fill-placeholders", "backup-restore"]
export const pathProperties = {
  docx: { type: "string", description: "Word 文件绝对路径（.docx/.docm）" },
  excel: { type: "string", description: "Excel 工作簿绝对路径（.xlsx/.xlsm）；留空使用默认同名工作簿" },
  paths: { type: "array", items: { type: "string" }, maxItems: 200, description: "批量处理的绝对路径清单" },
  templates: { type: "array", items: { type: "string" }, maxItems: 200, description: "Word 模板绝对路径清单" },
  items: {
    type: "array",
    maxItems: 200,
    description: "批量任务清单；模板生成使用 excel/template，批量删表使用 docx/pick/keepCaptionFor/reversible/excel",
    items: { type: "object" },
  },
  outDir: { type: "string", description: "输出目录绝对路径" },
  backupPath: { type: "string", description: "历史备份文件绝对路径" },
  target: { type: "string", description: "恢复目标文件绝对路径" },
}
export const runProperties = {
  ...pathProperties,
  dryRun: { type: "boolean", description: "仅预演，不写文件；Agent 处理写盘任务时应先传 true，用户确认后传 false" },
  force: { type: "boolean", description: "跳过文件占用检查；仅在用户明确同意时使用" },
  backup: { type: "boolean", description: "覆盖前备份到 _backup；Agent 处理写盘任务时建议保持 true" },
  noBackup: { type: "boolean", description: "兼容旧参数：覆盖前不备份" },
  withExcelLinks: { type: "boolean", description: "Excel 中写入定位 Word 链接" },
  totalsFormula: { type: "boolean", description: "合计行写 SUM 公式" },
  contents: {
    type: "object",
    properties: {
      tables: { type: "boolean" },
      marks: { type: "boolean" },
      placeholders: { type: "boolean" },
      keepHighlight: { type: "boolean" },
    },
    additionalProperties: false,
  },
  keepHighlight: { type: "boolean", description: "同步标注后保留高亮" },
  deleteEmptyRows: { type: "boolean", description: "删除 B 列标记为✕的行" },
  changelog: { type: "boolean", description: "导出差量改动清单" },
  pick: { type: "string", description: "候选/恢复编号，逗号分隔或 all" },
  keepCaption: { type: "boolean", description: "只删表格并保留标题" },
  keepCaptionFor: { type: "string", description: "保留标题的候选编号" },
  reversible: { type: "boolean", description: "可逆删除并保留 Excel 表块" },
  name: { type: "string", description: "新表名称" },
  nrows: { type: "integer", minimum: 1, maximum: 200 },
  ncols: { type: "integer", minimum: 1, maximum: 30 },
  after: { type: "string", description: "在指定表块后插入" },
  code: { type: "string", description: "表块编号" },
  column: { type: "string", description: "排序依据的列名" },
  descending: { type: "boolean", description: "降序；false 为升序" },
}
export const operationSpecs = {
  "blank-scan": { title: "扫描空白表", tool: "audit_query", required: ["docx"], optional: ["docx"], defaults: {}, dryRun: false, writes: false, example: { operation: "blank-scan", docx: "D:\\审计项目\\报告.docx" } },
  "deleted-list": { title: "查看可恢复表格", tool: "audit_query", required: ["docx"], optional: ["docx", "excel"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: false, example: { operation: "deleted-list", docx: "D:\\审计项目\\报告.docx", excel: "D:\\审计项目\\报告_表格数据.xlsx" } },
  "blank-batch-scan": { title: "批量扫描空白表", tool: "audit_query", required: ["paths"], optional: ["paths"], defaults: {}, dryRun: false, writes: false, example: { operation: "blank-batch-scan", paths: ["D:\\审计项目\\报告一.docx", "D:\\审计项目\\报告二.docx"] } },
  diff: { title: "差异比对", tool: "audit_query", required: ["docx"], optional: ["docx", "excel"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: false, example: { operation: "diff", docx: "D:\\审计项目\\报告.docx", excel: "D:\\审计项目\\报告_表格数据.xlsx" } },
  "table-list": { title: "查看工作簿表块", tool: "audit_query", required: ["docx"], optional: ["docx", "excel"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: false, example: { operation: "table-list", docx: "D:\\审计项目\\报告.docx", excel: "D:\\审计项目\\报告_表格数据.xlsx" } },
  "sort-list": { title: "查看排序设置", tool: "audit_query", required: ["docx"], optional: ["docx", "excel"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: false, example: { operation: "sort-list", docx: "D:\\审计项目\\报告.docx", excel: "D:\\审计项目\\报告_表格数据.xlsx" } },
  "template-scan": { title: "检查模板字段", tool: "audit_query", required: ["docx"], optional: ["docx"], defaults: {}, dryRun: false, writes: false, example: { operation: "template-scan", docx: "D:\\审计项目\\模板.docx" } },
  "backup-list": { title: "查看历史备份", tool: "audit_query", required: ["docx"], optional: ["docx", "excel"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: false, example: { operation: "backup-list", docx: "D:\\审计项目\\报告.docx", excel: "D:\\审计项目\\报告_表格数据.xlsx" } },
  export: { title: "创建 Excel", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force", "backup", "noBackup", "withExcelLinks", "totalsFormula"], defaults: { excel: "同名_表格数据.xlsx", withExcelLinks: true, totalsFormula: true, backup: true, force: false }, dryRun: true, writes: true, example: { operation: "export", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  writeback: { title: "同步到 Word", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force", "backup", "noBackup", "contents", "deleteEmptyRows", "changelog"], defaults: { excel: "同名_表格数据.xlsx", contents: { tables: true, marks: false, placeholders: false }, backup: true, force: false }, dryRun: true, writes: true, example: { operation: "writeback", docx: "D:\\审计项目\\报告.docx", excel: "D:\\审计项目\\报告_表格数据.xlsx", contents: { tables: true }, dryRun: true } },
  bookmark: { title: "重建或迁移书签", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force", "withExcelLinks"], defaults: { excel: "同名_表格数据.xlsx", withExcelLinks: true, force: false }, dryRun: true, writes: true, example: { operation: "bookmark", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  rejudge: { title: "重新判定标记", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force", "backup", "noBackup"], defaults: { excel: "同名_表格数据.xlsx", backup: true, force: false }, dryRun: true, writes: true, example: { operation: "rejudge", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  "clear-bookmarks": { title: "清理 Word 书签", tool: "audit_run", required: ["docx"], optional: ["docx", "dryRun", "force"], defaults: { force: false }, dryRun: true, writes: true, example: { operation: "clear-bookmarks", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  marks: { title: "导出或刷新标注", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force"], defaults: { excel: "同名_表格数据.xlsx", force: false }, dryRun: true, writes: true, example: { operation: "marks", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  "marks-sync": { title: "同步标注到 Word", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force", "keepHighlight"], defaults: { excel: "同名_表格数据.xlsx", keepHighlight: false, force: false }, dryRun: true, writes: true, example: { operation: "marks-sync", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  placeholders: { title: "补齐占位符工作表", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force"], defaults: { excel: "同名_表格数据.xlsx", force: false }, dryRun: true, writes: true, example: { operation: "placeholders", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  "placeholders-sync": { title: "同步占位符到 Word", tool: "audit_run", required: ["docx"], optional: ["docx", "excel", "dryRun", "force"], defaults: { excel: "同名_表格数据.xlsx", force: false }, dryRun: true, writes: true, example: { operation: "placeholders-sync", docx: "D:\\审计项目\\报告.docx", dryRun: true } },
  "blank-delete": { title: "删除空白表格", tool: "audit_run", required: ["docx", "pick"], optional: ["docx", "excel", "dryRun", "force", "pick", "keepCaption", "keepCaptionFor", "reversible"], defaults: { reversible: true, keepCaption: false, force: false }, dryRun: true, writes: true, example: { operation: "blank-delete", docx: "D:\\审计项目\\报告.docx", pick: "1", dryRun: true } },
  "restore-tables": { title: "恢复已删表格", tool: "audit_run", required: ["docx", "pick"], optional: ["docx", "excel", "dryRun", "force", "pick"], defaults: { excel: "同名_表格数据.xlsx", force: false }, dryRun: true, writes: true, example: { operation: "restore-tables", docx: "D:\\审计项目\\报告.docx", pick: "1", dryRun: true } },
  "blank-batch-delete": { title: "批量删除空白表", tool: "audit_run", required: ["items"], optional: ["items", "dryRun", "force"], defaults: { force: false }, dryRun: true, writes: true, example: { operation: "blank-batch-delete", items: [{ docx: "D:\\审计项目\\报告.docx", pick: "1", reversible: true }], dryRun: true } },
  "diff-export": { title: "导出差异清单", tool: "audit_run", required: ["docx"], optional: ["docx", "excel"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: true, example: { operation: "diff-export", docx: "D:\\审计项目\\报告.docx" } },
  "append-table": { title: "追加新表块", tool: "audit_run", required: ["docx", "name", "nrows", "ncols"], optional: ["docx", "excel", "name", "nrows", "ncols", "after"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: true, example: { operation: "append-table", docx: "D:\\审计项目\\报告.docx", name: "审计发现问题汇总", nrows: 3, ncols: 3 } },
  "sort-set": { title: "设置表格排序", tool: "audit_run", required: ["docx", "code", "column"], optional: ["docx", "excel", "code", "column", "descending"], defaults: { excel: "同名_表格数据.xlsx", descending: true }, dryRun: false, writes: true, example: { operation: "sort-set", docx: "D:\\审计项目\\报告.docx", code: "T02", column: "金额", descending: true } },
  "sort-clear": { title: "清除表格排序", tool: "audit_run", required: ["docx", "code"], optional: ["docx", "excel", "code"], defaults: { excel: "同名_表格数据.xlsx" }, dryRun: false, writes: true, example: { operation: "sort-clear", docx: "D:\\审计项目\\报告.docx", code: "T02" } },
  "batch-writeback": { title: "多报告批量同步", tool: "audit_run", required: ["paths"], optional: ["paths", "dryRun", "force", "backup", "noBackup", "contents", "keepHighlight"], defaults: { contents: { tables: true, marks: false, placeholders: false }, keepHighlight: true, backup: true, force: false }, dryRun: true, writes: true, example: { operation: "batch-writeback", paths: ["D:\\审计项目\\报告一.docx"], contents: { tables: true }, dryRun: true } },
  "template-batch": { title: "模板批量生成", tool: "audit_run", required: ["templates", "items", "outDir"], optional: ["templates", "items", "outDir", "dryRun", "contents", "keepHighlight"], defaults: { contents: { tables: true, marks: false, placeholders: true }, keepHighlight: true }, dryRun: true, writes: true, example: { operation: "template-batch", templates: ["D:\\审计项目\\模板.docx"], items: [{ excel: "D:\\审计项目\\底稿.xlsx", template: 1 }], outDir: "D:\\审计项目\\输出", dryRun: true } },
  "fill-placeholders": { title: "生成正文信息填写表", tool: "audit_run", required: ["docx", "paths"], optional: ["docx", "paths"], defaults: {}, dryRun: false, writes: true, example: { operation: "fill-placeholders", docx: "D:\\审计项目\\模板.docx", paths: ["D:\\审计项目\\底稿.xlsx"] } },
  "backup-restore": { title: "恢复历史备份", tool: "audit_run", required: ["backupPath", "target"], optional: ["backupPath", "target"], defaults: {}, dryRun: false, writes: true, example: { operation: "backup-restore", backupPath: "D:\\审计项目\\_backup\\报告.bak.docx", target: "D:\\审计项目\\报告.docx" } },
}

const isObject = value => Boolean(value) && typeof value === "object" && !Array.isArray(value)

function validatePath(value, label, suffixes = []) {
  if (typeof value !== "string" || !value.trim()) throw new Error(`${label}必须使用绝对路径`)
  if (/[\u0000-\u001f]/.test(value)) throw new Error(`${label}含有控制字符`)
  if (!path.isAbsolute(value)) throw new Error(`${label}必须使用绝对路径`)
  if (suffixes.length && !suffixes.some(suffix => value.toLowerCase().endsWith(suffix))) {
    throw new Error(`${label}只支持 ${suffixes.join("/")} 文件`)
  }
  return value
}

export function validateRunArgs(args, allowedOperations) {
  if (!isObject(args)) throw new Error("参数必须是对象")
  if (!allowedOperations.includes(args.operation)) throw new Error("不支持的操作")
  const spec = operationSpecs[args.operation]
  if (!spec) throw new Error("该操作缺少能力定义")
  if (args.dryRun !== undefined && !spec.dryRun) throw new Error(`${spec.title}不支持预演，请先用相关只读查询检查。`)
  const allowedKeys = new Set(["operation", ...spec.required, ...spec.optional])
  for (const key of Object.keys(args)) {
    if (!allowedKeys.has(key)) throw new Error(`${spec.title}不支持参数：${key}`)
  }
  const labels = { docx: "Word 文件", excel: "Excel 工作簿", paths: "文件清单", templates: "模板清单", items: "任务清单", outDir: "输出目录", pick: "候选编号", name: "新表名称", nrows: "行数", ncols: "列数", code: "表块编号", column: "排序列名", backupPath: "历史备份文件", target: "恢复目标文件" }
  for (const key of spec.required) {
    const value = args[key]
    if (value === undefined || value === null || value === "" || (Array.isArray(value) && value.length === 0)) {
      throw new Error(`${spec.title}缺少${labels[key] || key}。`)
    }
  }
  const request = { operation: args.operation }
  if (args.docx !== undefined && args.docx !== "") request.docx = validatePath(args.docx, "Word 文件", [".docx", ".docm"])
  if (args.excel !== undefined && args.excel !== "") request.excel = validatePath(args.excel, "Excel 工作簿", [".xlsx", ".xlsm"])
  for (const key of ["paths", "templates"]) {
    if (args[key] === undefined) continue
    if (!Array.isArray(args[key]) || args[key].length > 200) throw new Error(`${key} 必须是路径数组，且最多 200 项`)
    const suffixes = key === "templates" || args.operation !== "fill-placeholders" ? [".docx", ".docm"] : [".xlsx", ".xlsm"]
    request[key] = args[key].map((value, index) => validatePath(value, `${key}[${index}]`, suffixes))
  }
  if (args.items !== undefined) {
    if (!Array.isArray(args.items) || args.items.length > 200 || args.items.some(item => !isObject(item))) {
      throw new Error("items 必须是对象数组，且最多 200 项")
    }
    if (args.operation === "template-batch") {
      request.items = args.items.map((item, index) => {
        const allowed = new Set(["excel", "template"])
        for (const key of Object.keys(item)) if (!allowed.has(key)) throw new Error(`items[${index}] 不支持参数：${key}`)
        if (!Number.isInteger(item.template) || item.template < 1 || item.template > request.templates.length) throw new Error(`items[${index}].template 必须是 1～${request.templates.length} 的模板编号`)
        return { excel: validatePath(item.excel, `items[${index}].excel`, [".xlsx", ".xlsm"]), template: item.template }
      })
    } else if (args.operation === "blank-batch-delete") {
      request.items = args.items.map((item, index) => {
        const allowed = new Set(["docx", "pick", "keepCaptionFor", "reversible", "excel"])
        for (const key of Object.keys(item)) if (!allowed.has(key)) throw new Error(`items[${index}] 不支持参数：${key}`)
        if (typeof item.pick !== "string" || !item.pick.trim()) throw new Error(`items[${index}].pick 必须填写候选编号`)
        if (item.reversible !== undefined && typeof item.reversible !== "boolean") throw new Error(`items[${index}].reversible 必须是布尔值`)
        const value = { docx: validatePath(item.docx, `items[${index}].docx`, [".docx", ".docm"]), pick: item.pick.trim() }
        if (item.keepCaptionFor !== undefined) {
          if (typeof item.keepCaptionFor !== "string") throw new Error(`items[${index}].keepCaptionFor 必须是文本`)
          value.keepCaptionFor = item.keepCaptionFor
        }
        if (item.reversible !== undefined) value.reversible = item.reversible
        if (item.excel !== undefined && item.excel !== "") value.excel = validatePath(item.excel, `items[${index}].excel`, [".xlsx", ".xlsm"])
        return value
      })
    } else {
      request.items = args.items
    }
  }
  if (args.outDir !== undefined) request.outDir = validatePath(args.outDir, "输出目录")
  if (args.backupPath !== undefined) request.backupPath = validatePath(args.backupPath, "备份文件", [".docx", ".docm", ".xlsx", ".xlsm"])
  if (args.target !== undefined) request.target = validatePath(args.target, "目标文件", [".docx", ".docm", ".xlsx", ".xlsm"])
  if (request.backupPath && request.target && path.extname(request.backupPath).toLowerCase() !== path.extname(request.target).toLowerCase()) {
    throw new Error("备份文件与目标文件的扩展名必须一致。")
  }
  for (const key of ["dryRun", "force", "backup", "noBackup", "withExcelLinks", "totalsFormula", "keepHighlight", "deleteEmptyRows", "changelog", "keepCaption", "reversible", "descending"]) {
    if (args[key] !== undefined) {
      if (typeof args[key] !== "boolean") throw new Error(`${key} 必须是布尔值`)
      request[key] = args[key]
    }
  }
  for (const key of ["pick", "keepCaptionFor", "name", "after", "code", "column"]) {
    if (args[key] !== undefined) {
      if (typeof args[key] !== "string") throw new Error(`${key} 必须是文本`)
      request[key] = args[key]
    }
  }
  for (const key of ["nrows", "ncols"]) {
    if (args[key] !== undefined) {
      if (!Number.isInteger(args[key])) throw new Error(`${key} 必须是整数`)
      const maximum = key === "nrows" ? 200 : 30
      if (args[key] < 1 || args[key] > maximum) throw new Error(`${key} 必须在 1～${maximum} 之间`)
      request[key] = args[key]
    }
  }
  if (args.contents !== undefined) {
    if (!isObject(args.contents)) throw new Error("同步内容格式不正确")
    request.contents = {}
    for (const key of ["tables", "marks", "placeholders", "keepHighlight"]) {
      if (args.contents[key] !== undefined) {
        if (typeof args.contents[key] !== "boolean") throw new Error(`contents.${key} 必须是布尔值`)
        request.contents[key] = args.contents[key]
      }
    }
    if (!["tables", "marks", "placeholders"].some(key => request.contents[key] === true)) throw new Error("同步内容至少勾选一项。")
  }
  return request
}
