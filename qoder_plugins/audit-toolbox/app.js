(() => {
  const protocolVersion = "2026-01-26"
  const groups = ["常用", "内容", "维护", "批量", "备份"]
  const contentOptions = [
    { key: "tables", label: "表格数据", hint: "按 Excel 内容回写 Word 表格" },
    { key: "marks", label: "链接标注", hint: "更新高亮内容及其定位" },
    { key: "placeholders", label: "正文占位符", hint: "填入报告名称、日期等信息" },
  ]
  const operations = {
    export: { group: "常用", title: "创建 Excel", description: "从 Word 提取表格、标注和占位符。三类内容只需存在一种即可导出，三个工作表始终创建。", requires: ["docx", "excel"], dryRun: true, action: "创建工作簿", fields: [
      { key: "withExcelLinks", type: "checkbox", label: "Excel 中写入 Word 定位链接", default: true, hint: "点击定位链接可选中 Word 中的对应表格。" },
      { key: "totalsFormula", type: "checkbox", label: "合计行写入 SUM 公式", default: true },
      { key: "backup", type: "checkbox", label: "覆盖前备份原工作簿", default: true, hint: "备份保存在工作簿旁的 _backup 文件夹。" },
      { key: "force", type: "checkbox", label: "文件被占用时仍尝试保存", default: false, danger: true },
    ] },
    writeback: { group: "常用", title: "同步到 Word", description: "把工作簿中的改动写回报告。默认只同步表格，可同时处理标注和正文占位符。", requires: ["docx", "excel"], dryRun: true, action: "同步写盘", fields: [
      { key: "contents", type: "contents", label: "同步哪些内容", defaults: { tables: true } },
      { key: "deleteEmptyRows", type: "checkbox", label: "删除 B 列标记为 ✕ 的数据行", default: false, hint: "仅影响同步的表格数据。" },
      { key: "changelog", type: "checkbox", label: "导出逐格改动清单", default: false },
      { key: "backup", type: "checkbox", label: "写盘前备份 Word 和 Excel", default: true },
      { key: "force", type: "checkbox", label: "文件被占用时仍尝试保存", default: false, danger: true },
    ] },
    diff: { group: "常用", title: "差异比对", description: "只读比较 Excel 与 Word 的差异，逐格列出旧值和新值，不修改任何文件。", query: true, requires: ["docx", "excel"], action: "比较差异" },
    "diff-export": { group: "常用", title: "导出差异清单", description: "把当前 Excel 与 Word 的差异导出为工作簿旁的改动清单，不修改 Word 内容。", requires: ["docx", "excel"], action: "导出差异清单" },
    rejudge: { group: "常用", title: "重新判定 ✓ / ✕", description: "按 Excel 当前内容重新判定表头、✓、✕ 标记，不改 Word。", requires: ["docx", "excel"], dryRun: true, action: "重新判定", fields: [{ key: "backup", type: "checkbox", label: "保存前备份工作簿", default: true }] },
    marks: { group: "内容", title: "导出 / 刷新标注", description: "把 Word 中的高亮标注写入「链接标注」工作表，Word 文字保持不变。", requires: ["docx", "excel"], dryRun: true, action: "刷新标注工作表" },
    "marks-sync": { group: "内容", title: "同步标注到 Word", description: "将「链接标注」工作表中的更新写回 Word，可按需移除已替换的高亮。", requires: ["docx", "excel"], dryRun: true, action: "同步标注", fields: [{ key: "keepHighlight", type: "checkbox", label: "保留 Word 中的高亮外观", default: false }] },
    placeholders: { group: "内容", title: "补齐占位符工作表", description: "扫描正文、页眉页脚的 {字段名}，生成或补齐「正文信息」工作表并保留已填值。", requires: ["docx", "excel"], dryRun: true, action: "补齐占位符" },
    "placeholders-sync": { group: "内容", title: "同步占位符到 Word", description: "将「正文信息」中的报告名称、日期、数字等值填入 Word，保留原有字体和版式。", requires: ["docx", "excel"], dryRun: true, action: "同步占位符" },
    "template-scan": { group: "内容", title: "检查模板字段", description: "只读列出模板正文及页眉页脚中的占位符字段，便于确认批量生成前的填写表。", query: true, requires: ["docx"], action: "检查模板" },
    bookmark: { group: "维护", title: "重建 / 迁移书签", description: "把 Word 表格重新锚定为整表书签，并清理正文遗留的〔→Excel〕文字，不改 Excel 数据。", requires: ["docx"], dryRun: true, action: "重建书签", fields: [{ key: "withExcelLinks", type: "checkbox", label: "沿用已有 Excel 定位链接", default: true }] },
    "clear-bookmarks": { group: "维护", title: "清理 Word 书签", description: "清除工具写入的定位书签。执行后需重新导出 Excel 才能继续正常同步。", requires: ["docx"], dryRun: true, action: "清理书签", danger: true },
    "table-list": { group: "维护", title: "查看工作簿表块", description: "只读查看表块编号、表名、书签和行列数，也可从这里选择新表插入位置。", query: true, requires: ["docx", "excel"], action: "查看表块" },
    "incremental-export": { group: "维护", title: "增量导出", description: "把 Word 中尚未登记的新表格、高亮标注和占位符追加到现有工作簿；既有行、编号与公式位置保持不变。", requires: ["docx", "excel"], dryRun: true, action: "增量导出", fields: [{ key: "backup", type: "checkbox", label: "保存前备份 Word 与 Excel", default: true }] },
    "sort-list": { group: "维护", title: "查看排序设置", description: "只读查看各表块当前排序及可选金额列。排序在同步到 Word 时应用。", query: true, requires: ["docx", "excel"], action: "查看排序" },
    "sort-set": { group: "维护", title: "设置表格排序", description: "按列名设置表块排序；表头与合计行保持不动，下一次同步时应用。", requires: ["docx", "excel"], action: "保存排序设置", fields: [
      { key: "code", type: "text", label: "表块编号", placeholder: "例如 T02", required: true },
      { key: "column", type: "text", label: "排序列名", placeholder: "填写表头文字", required: true },
      { key: "descending", type: "checkbox", label: "降序（金额从大到小）", default: true },
    ] },
    "sort-clear": { group: "维护", title: "清除表格排序", description: "移除指定表块的排序设置，不改表格数据。", requires: ["docx", "excel"], action: "清除排序", fields: [{ key: "code", type: "text", label: "表块编号", placeholder: "例如 T02", required: true }] },
    "blank-scan": { group: "维护", title: "扫描空白表格", description: "只读找出整表全空或仅表头有字的表格，扫描后可直接勾选删除。", query: true, requires: ["docx"], action: "扫描空白表" },
    "blank-delete": { group: "维护", title: "删除空白表格", description: "按候选编号删除表格；可连表标题一起删除。建议先扫描，避免误删。", requires: ["docx"], dryRun: true, action: "删除空白表", danger: true, fields: [
      { key: "pick", type: "text", label: "候选编号", placeholder: "例如 1,3；全部填写 all", required: true, hint: "编号来自「扫描空白表格」结果。" },
      { key: "excel", type: "path", label: "对应 Excel 工作簿（可逆删除时必填）", suffix: "xlsx", optional: true },
      { key: "keepCaption", type: "checkbox", label: "只删表格，保留所有表标题", default: false },
      { key: "keepCaptionFor", type: "text", label: "只保留这些候选的标题", placeholder: "例如 2,4；可留空" },
      { key: "reversible", type: "checkbox", label: "可逆删除，保留 Excel 表块供恢复", default: true },
      { key: "force", type: "checkbox", label: "文件被占用时仍尝试保存", default: false, danger: true },
    ] },
    "deleted-list": { group: "维护", title: "查看可恢复表格", description: "列出工作簿中标记为「已删除」的表块，扫描后可直接勾选恢复到原位置。", query: true, requires: ["docx", "excel"], action: "查看已删表块" },
    "restore-tables": { group: "维护", title: "恢复已删表格", description: "按恢复清单编号重建表格与书签，并恢复到原位置；锚点缺失时退到文档末尾。", requires: ["docx", "excel"], dryRun: true, action: "恢复表格", fields: [
      { key: "pick", type: "text", label: "恢复编号", placeholder: "例如 1,3；全部填写 all", required: true },
      { key: "force", type: "checkbox", label: "文件被占用时仍尝试保存", default: false, danger: true },
    ] },
    "batch-writeback": { group: "批量", title: "多报告批量同步", description: "每份 Word 配同名工作簿逐个同步，单份失败不影响其他报告。", paths: { key: "paths", label: "Word 报告清单", suffix: "docx" }, dryRun: true, action: "批量同步", fields: [
      { key: "contents", type: "contents", label: "批量同步内容", defaults: { tables: true } },
      { key: "keepHighlight", type: "checkbox", label: "保留标注高亮", default: true },
      { key: "backup", type: "checkbox", label: "每份写盘前备份", default: true },
      { key: "force", type: "checkbox", label: "文件被占用时仍尝试保存", default: false, danger: true },
    ] },
    "blank-batch-scan": { group: "批量", title: "批量扫描空白表", description: "一次扫描多份报告，按报告分组显示候选空白表；可勾选后进入批量删除。", query: true, paths: { key: "paths", label: "Word 报告清单", suffix: "docx" }, action: "批量扫描" },
    "blank-batch-delete": { group: "批量", title: "批量删除空白表", description: "按扫描结果逐份删除已勾选的空白表，单份失败不影响其他报告。", special: "blank-batch", dryRun: true, action: "批量删除", danger: true, fields: [
      { key: "reversible", type: "checkbox", label: "默认可逆删除，保留 Excel 表块", default: true },
      { key: "force", type: "checkbox", label: "文件被占用时仍尝试保存", default: false, danger: true },
    ] },
    "template-batch": { group: "批量", title: "多模板批量生成", description: "为每个 Excel 指定所属 Word 模板，在模板副本上生成新报告，不改模板原件和源工作簿。", special: "template-batch", dryRun: true, action: "批量生成报告", fields: [
      { key: "templates", type: "paths", label: "Word 模板清单", suffix: "docx", required: true },
      { key: "outDir", type: "outdir", label: "报告输出目录", required: true },
      { key: "contents", type: "contents", label: "批量生成内容", defaults: { tables: true, placeholders: true } },
      { key: "keepHighlight", type: "checkbox", label: "保留标注高亮", default: true },
    ] },
    "fill-placeholders": { group: "批量", title: "生成正文信息填写表", description: "按模板字段为多份 Excel 补齐「正文信息」工作表，保留已经填写的值。", requires: ["docx"], paths: { key: "paths", label: "Excel 工作簿清单", suffix: "xlsx" }, action: "补齐填写表" },
    "backup-list": { group: "备份", title: "查看历史备份", description: "只读列出 Word 与 Excel 的历史备份，可直接选择某个备份恢复到当前文件。", query: true, requires: ["docx", "excel"], action: "查看备份" },
    "backup-restore": { group: "备份", title: "恢复历史备份", description: "用历史备份覆盖目标文件。恢复前会先备份目标现状，可再次回滚。", requires: [], action: "恢复备份", danger: true, fields: [
      { key: "backupPath", type: "path", label: "历史备份文件", suffix: "any", required: true },
      { key: "target", type: "path", label: "恢复到的目标文件", suffix: "any", required: true, hint: "备份与目标必须同为 Word 或同为 Excel，扩展名一致。" },
    ] },
  }
  const state = { initialized: false, capabilities: {}, operation: "export", engine: null, busy: false, taskId: null, task: null, logs: [], pendingAction: null, values: {}, templateItems: [], templateScans: {}, batchItems: [], excelTouched: false, renderedResult: null }
  const pending = new Map(), hostPending = new Map(), timers = new Map()
  let requestId = 0
  const $ = id => document.getElementById(id)
  const el = (tag, className, text) => { const node = document.createElement(tag); if (className) node.className = className; if (text !== undefined) node.textContent = String(text); return node }
  const listValue = value => Array.isArray(value) ? value : String(value || "").split(/\r?\n/).map(item => item.trim()).filter(Boolean)
  const isAbsolute = value => /^(?:[A-Za-z]:[\\/]|\\\\)/.test(value)
  const dirname = value => value.replace(/[\\/][^\\/]+$/, "")
  const basename = value => value.split(/[\\/]/).pop()
  const current = () => operations[state.operation]
  const fieldId = key => `field-${key}`
  const valueKey = (operation, key) => `${operation}.${key}`
  const getValue = (key, fallback = "") => state.values[valueKey(state.operation, key)] ?? fallback
  const setValue = (operation, key, value) => { state.values[valueKey(operation, key)] = value }
  const post = message => window.parent.postMessage(message, "*")

  function rpc(method, params, target, timeoutMs = 20000) {
    const id = `app-${++requestId}`
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { target.delete(id); reject(new Error(`请求超时：${method}`)) }, timeoutMs)
      target.set(id, { resolve, reject, timer })
      try { post({ jsonrpc: "2.0", id, method, params }) } catch { clearTimeout(timer); target.delete(id); reject(new Error("无法连接 Qoder，请重新打开页面。")) }
    })
  }
  async function callTool(name, args = {}) {
    const response = await rpc("tools/call", { name, arguments: args }, pending)
    if (response?.error) throw new Error(response.error.message || "工具调用失败")
    return response?.result ?? { content: [], structuredContent: {} }
  }
  function renderNav() {
    const nav = $("task-nav")
    nav.replaceChildren()
    for (const group of groups) {
      const section = el("div", "nav-group")
      section.appendChild(el("p", "nav-title", group))
      for (const [id, op] of Object.entries(operations).filter(([, item]) => item.group === group)) {
        const button = el("button", "nav-button", `${op.title}`)
        button.type = "button"
        button.dataset.operation = id
        button.setAttribute("aria-current", String(id === state.operation))
        button.appendChild(el("span", "nav-icon", String(Object.keys(operations).indexOf(id) + 1)))
        button.addEventListener("click", () => switchOperation(id))
        section.appendChild(button)
      }
      nav.appendChild(section)
    }
  }
  function switchOperation(operation, seed = {}) {
    if (state.busy) return
    for (const [key, value] of Object.entries(seed)) setValue(operation, key, value)
    state.operation = operation
    state.pendingAction = null
    $("confirmation").hidden = true
    $("result-card").hidden = true
    state.renderedResult = null
    renderNav()
    renderOperation()
  }
  function renderOperation() {
    const op = current()
    $("page-group").textContent = op.group
    $("page-title").textContent = op.title
    $("page-description").textContent = op.description
    $("page-badge").textContent = op.query ? "只读 · 不改文件" : op.danger ? "写盘 · 需要确认" : "本地处理"
    $("dry-run-label").hidden = !op.dryRun
    $("dry-run").checked = false
    renderFields(op)
    updateRunButton()
  }
  function renderFields(op) {
    const root = $("form-fields")
    root.replaceChildren()
    let step = 0
    if (op.requires?.length || op.paths || op.special) {
      root.appendChild(sectionTitle(++step, "选择文件", "只读取你明确选择的本地文件"))
      if (op.requires?.includes("docx")) root.appendChild(pathField("docx", "Word 报告 / 模板", "docx", false))
      if (op.requires?.includes("excel")) root.appendChild(pathField("excel", "Excel 工作簿", "xlsx", true))
      if (op.paths) root.appendChild(pathsField(op.paths))
      if (op.special === "template-batch") root.appendChild(templateMappingFields())
      if (op.special === "blank-batch") root.appendChild(batchSelectionFields())
    }
    if (op.fields?.length) {
      root.appendChild(sectionTitle(++step, "操作选项", op.query ? "确认查询条件" : "先预演，确认结果后再正式执行"))
      for (const field of op.fields) root.appendChild(renderField(field))
    }
  }
  function sectionTitle(step, title, hint) {
    const section = el("div", "form-section")
    const heading = el("h3", "section-title")
    heading.appendChild(el("span", "step", step))
    heading.appendChild(el("span", "", title))
    if (hint) heading.appendChild(el("small", "", hint))
    section.appendChild(heading)
    return section
  }
  function pathField(key, label, suffix, optional = false) {
    const wrap = el("label", "field")
    wrap.appendChild(fieldLabel(label, optional ? "留空使用默认同名工作簿" : "Windows 绝对路径"))
    const row = el("div", "path-row")
    const input = el("input", "text-input")
    input.id = fieldId(key); input.type = "text"; input.placeholder = suffix === "docx" ? "C:\\报告\\审计报告.docx" : "C:\\报告\\审计报告_表格数据.xlsx"
    input.value = getValue(key, "")
    input.autocomplete = "off"; input.spellcheck = false
    input.addEventListener("input", () => {
      setValue(state.operation, key, input.value)
      if (key === "docx" && current().requires?.includes("excel") && !state.excelTouched) {
        const excel = $("field-excel")
        if (excel) { excel.value = defaultExcel(input.value); setValue(state.operation, "excel", excel.value) }
      }
    })
    const pick = el("button", "pick-button", "选择文件…")
    pick.type = "button"
    pick.dataset.pick = key
    pick.dataset.suffix = suffix
    pick.addEventListener("click", () => pickFiles(key, suffix))
    row.append(input, pick)
    wrap.appendChild(row)
    return wrap
  }
  function defaultExcel(docx) { return docx.replace(/\.(docx|docm)$/i, "") + "_表格数据.xlsx" }
  function fieldLabel(label, hint) {
    const node = el("div", "field-label")
    node.appendChild(el("span", "", label))
    if (hint) node.appendChild(el("small", "", hint))
    return node
  }
  function pathsField(config) {
    const wrap = el("label", "field")
    wrap.appendChild(fieldLabel(config.label, "每行一个路径，最多 200 份"))
    const textarea = el("textarea")
    textarea.id = fieldId(config.key)
    textarea.value = getValue(config.key, "").join("\n")
    textarea.placeholder = config.suffix === "docx" ? "C:\\报告\\2025年度审计报告.docx" : "C:\\数据\\2025年度审计报告_表格数据.xlsx"
    textarea.addEventListener("input", () => setValue(state.operation, config.key, listValue(textarea.value)))
    const row = el("div", "path-row")
    row.style.marginTop = "8px"
    const pick = el("button", "pick-button", "选择多个文件…")
    pick.type = "button"
    pick.addEventListener("click", () => pickPaths(config.key, config.suffix))
    const hint = el("span", "field-hint", "也可直接粘贴完整路径")
    row.append(pick, hint)
    wrap.append(textarea, row)
    return wrap
  }
  function renderField(field) {
    let wrap = el("label", "field")
    wrap.appendChild(fieldLabel(field.label, field.required ? "必填" : ""))
    let control
    if (field.type === "contents") {
      wrap = el("div", "field")
      wrap.appendChild(fieldLabel(field.label, "至少选择一项"))
      control = el("div", "check-grid")
      const saved = getValue(field.key, null)
      for (const item of contentOptions) {
        const row = el("label", "check-row")
        const input = el("input"); input.type = "checkbox"; input.id = `content-${item.key}`
        input.checked = saved ? Boolean(saved[item.key]) : Boolean(field.defaults[item.key])
        const text = el("span", "", item.label); text.appendChild(el("small", "", item.hint))
        row.append(input, text)
        input.addEventListener("change", () => {
          const values = { keepHighlight: $("content-keepHighlight")?.checked ?? true }
          for (const option of contentOptions) values[option.key] = $(`content-${option.key}`).checked
          setValue(state.operation, field.key, values)
        })
        control.appendChild(row)
      }
      if (!opUsesTopLevelHighlight()) {
        const row = el("label", "check-row")
        const input = el("input"); input.type = "checkbox"; input.id = "content-keepHighlight"; input.checked = saved?.keepHighlight ?? true
        const text = el("span", "", "保留标注高亮"); text.appendChild(el("small", "", "不改变 Word 中已高亮的文字外观"))
        row.append(input, text); control.appendChild(row)
        input.addEventListener("change", () => {
          const values = {}
          for (const option of contentOptions) values[option.key] = $(`content-${option.key}`).checked
          values.keepHighlight = input.checked
          setValue(state.operation, field.key, values)
        })
      }
    } else if (field.type === "checkbox") {
      const row = el("label", "check-row")
      control = el("input"); control.type = "checkbox"; control.id = fieldId(field.key); control.checked = getValue(field.key, field.default)
      control.style.width = "18px"; control.style.height = "18px"; control.style.accentColor = "var(--accent)"
      const text = el("span", "", field.label)
      if (field.hint) text.appendChild(el("small", "", field.hint))
      row.append(control, text)
      control.addEventListener("change", () => setValue(state.operation, field.key, control.checked))
      return row
    } else if (field.type === "number") {
      control = el("input", "text-input"); control.type = "number"; control.min = field.min; control.max = field.max; control.value = getValue(field.key, field.default)
    } else if (field.type === "path" || field.type === "outdir") {
      const row = el("div", "path-row")
      control = el("input", "text-input"); control.placeholder = field.type === "outdir" ? "C:\\报告\\生成结果" : "C:\\报告\\文件.xlsx"
      control.value = getValue(field.key, "")
      const pick = el("button", "pick-button", field.type === "outdir" ? "使用模板目录" : "选择文件…")
      pick.type = "button"
      pick.addEventListener("click", () => {
        if (field.type === "outdir") {
          const templates = getValue("templates", [])
          if (!templates.length) return showError("请先选择至少一个 Word 模板。")
          control.value = dirname(templates[0]); control.dispatchEvent(new Event("input"))
        } else pickFiles(field.key, field.suffix === "any" ? "mixed" : field.suffix)
      })
      row.append(control, pick)
      wrap.appendChild(row)
    } else {
      control = el("input", "text-input"); control.type = "text"; control.placeholder = field.placeholder || ""
      control.value = getValue(field.key, "")
    }
    control.id = fieldId(field.key)
    if (field.type !== "path" && field.type !== "outdir") wrap.appendChild(control)
    control.addEventListener("input", () => {
      setValue(state.operation, field.key, field.type === "number" ? Number(control.value) : control.value)
      if (field.key === "templates") refreshTemplateSelects()
    })
    if (field.hint) wrap.appendChild(el("p", "field-hint", field.hint))
    return wrap
  }
  function opUsesTopLevelHighlight() { return ["batch-writeback", "template-batch", "marks-sync"].includes(state.operation) }
  function templateMappingFields() {
    const box = el("div", "field")
    box.appendChild(fieldLabel("Excel 与模板对应关系", "每份 Excel 必须指定所属模板"))
    const tools = el("div", "result-actions")
    const add = el("button", "secondary-button", "添加 Excel 文件…")
    add.type = "button"; add.addEventListener("click", () => pickTemplateItems())
    const scan = el("button", "secondary-button", "检查所选模板")
    scan.type = "button"; scan.addEventListener("click", scanTemplates)
    tools.append(add, scan)
    const list = el("div", "mapping-list")
    const renderRows = () => {
      list.replaceChildren()
      const templates = getValue("templates", [])
      if (!state.templateItems.length) list.appendChild(el("p", "empty", "还没有添加 Excel 工作簿。"))
      state.templateItems.forEach((item, index) => {
        const row = el("div", "mapping-row")
        const path = el("input", "text-input"); path.value = item.excel; path.title = item.excel
        path.addEventListener("input", () => { state.templateItems[index].excel = path.value })
        const select = el("select"); select.setAttribute("aria-label", `第 ${index + 1} 份工作簿所属模板`)
        for (let i = 0; i < templates.length; i++) { const option = el("option", "", `模板 ${i + 1} · ${basename(templates[i])}`); option.value = String(i + 1); select.appendChild(option) }
        select.value = String(item.template || 1)
        select.addEventListener("change", () => { state.templateItems[index].template = Number(select.value) })
        const remove = el("button", "quiet-button", "移除"); remove.type = "button"; remove.addEventListener("click", () => { state.templateItems.splice(index, 1); renderRows() })
        const actions = el("div"); actions.style.display = "flex"; actions.style.gap = "6px"; actions.append(select, remove)
        row.append(path, actions); list.appendChild(row)
      })
    }
    renderRows()
    box.append(tools, list)
    box._renderRows = renderRows
    return box
  }
  function batchSelectionFields() {
    const box = el("div", "field")
    box.appendChild(fieldLabel("本次删除清单", "从批量扫描结果带入"))
    if (!state.batchItems.length) {
      box.appendChild(el("div", "notice", "尚未选择空白表。请先到「批量扫描空白表」勾选，再进入本页。"))
      const scan = el("button", "secondary-button", "先去批量扫描")
      scan.type = "button"; scan.style.marginTop = "10px"; scan.addEventListener("click", () => switchOperation("blank-batch-scan"))
      box.appendChild(scan)
    } else {
      const count = state.batchItems.reduce((sum, item) => sum + item.pick.split(",").filter(Boolean).length, 0)
      box.appendChild(el("div", "notice", `已选择 ${state.batchItems.length} 份报告中的 ${count} 张空白表。`))
      const list = el("div", "choice-list")
      for (const item of state.batchItems) list.appendChild(el("div", "choice-row", `${basename(item.docx)}：候选 ${item.pick}${item.keepCaptionFor ? `（保留标题 ${item.keepCaptionFor}）` : ""}`))
      list.style.marginTop = "10px"; box.appendChild(list)
    }
    return box
  }
  async function pickFiles(target, suffix) {
    if (!state.capabilities.experimental?.["qoder/files"]?.pick) return showError("当前 Qoder 版本不支持文件选择器，请直接粘贴完整路径。")
    try {
      const extensions = suffix === "mixed" ? ["docx", "docm", "xlsx", "xlsm"] : [suffix]
      const response = await rpc("qoder/files/pick", { extensions }, hostPending)
      if (response?.error) throw new Error(response.error.message || "文件选择失败")
      const file = response?.result?.files?.[0]
      if (!file?.path) return
      const input = $(fieldId(target)); if (input) { input.value = file.path; input.dispatchEvent(new Event("input")) }
      setValue(state.operation, target, file.path)
      if (target === "excel") state.excelTouched = true
      if (target === "docx" && current().requires?.includes("excel") && !state.excelTouched) {
        const excel = $("field-excel"); if (excel) { excel.value = defaultExcel(file.path); setValue(state.operation, "excel", excel.value) }
      }
    } catch (error) { showError(error.message) }
  }
  async function pickPaths(target, suffix) {
    if (!state.capabilities.experimental?.["qoder/files"]?.pick) return showError("当前 Qoder 版本不支持文件选择器，请逐行粘贴完整路径。")
    try {
      const response = await rpc("qoder/files/pick", { extensions: [suffix] }, hostPending)
      if (response?.error) throw new Error(response.error.message || "文件选择失败")
      const paths = (response?.result?.files || []).map(file => file.path).filter(Boolean)
      if (!paths.length) return
      const input = $(fieldId(target))
      if (input) { input.value = paths.join("\n"); input.dispatchEvent(new Event("input")) }
      setValue(state.operation, target, paths)
      if (target === "templates") refreshTemplateSelects()
    } catch (error) { showError(error.message) }
  }
  async function pickTemplateItems() {
    if (!state.capabilities.experimental?.["qoder/files"]?.pick) return showError("当前 Qoder 版本不支持文件选择器，请在对应关系中逐项填写 Excel 完整路径。")
    try {
      const response = await rpc("qoder/files/pick", { extensions: ["xlsx", "xlsm"] }, hostPending)
      if (response?.error) throw new Error(response.error.message || "文件选择失败")
      const paths = (response?.result?.files || []).map(file => file.path).filter(Boolean)
      if (paths.length) { paths.forEach(excel => state.templateItems.push({ excel, template: 1 })); refreshTemplateSelects() }
    } catch (error) { showError(error.message) }
  }
  function refreshTemplateSelects() {
    const fields = [...$("form-fields").children]
    for (const field of fields) if (field._renderRows) field._renderRows()
  }
  function collectArgs() {
    const op = current()
    const args = { operation: state.operation }
    const validators = { docx: /\.(docx|docm)$/i, excel: /\.(xlsx|xlsm)$/i }
    for (const key of op.requires || []) {
      const value = String(getValue(key, "")).trim()
      if (!value) { if (key === "excel") continue; throw new Error(`请选择 Word 报告 / 模板。`) }
      validateAbsolute(value, key === "docx" ? "Word 文件" : "Excel 工作簿", validators[key])
      args[key] = value
    }
    if (op.paths) {
      const paths = listValue(getValue(op.paths.key, []))
      if (!paths.length) throw new Error(`请至少选择一个${op.paths.label}文件。`)
      const suffix = op.paths.suffix === "docx" ? /\.(docx|docm)$/i : /\.(xlsx|xlsm)$/i
      paths.forEach(value => validateAbsolute(value, op.paths.label, suffix))
      args[op.paths.key] = paths
    }
    for (const field of op.fields || []) {
      if (["templates", "outDir"].includes(field.key) || (field.key === "excel" && op.requires?.includes("excel"))) continue
      const value = getValue(field.key, field.default)
      if (field.type === "checkbox") args[field.key] = Boolean(value)
      else if (field.type === "contents") {
        const values = { ...field.defaults, ...(value || {}) }
        if (!values.tables && !values.marks && !values.placeholders) throw new Error("同步内容至少勾选一项。")
        args.contents = values
      } else if (field.type === "number") {
        if (!Number.isInteger(value) || value < field.min || value > field.max) throw new Error(`${field.label}必须是 ${field.min}～${field.max} 的整数。`)
        args[field.key] = value
      } else if (String(value ?? "").trim()) args[field.key] = String(value).trim()
      else if (field.required) throw new Error(`请填写${field.label}。`)
    }
    if (op.special === "template-batch") {
      const templates = listValue(getValue("templates", []))
      const outDir = String(getValue("outDir", "")).trim()
      if (!templates.length) throw new Error("请至少选择一个 Word 模板。")
      if (!outDir) throw new Error("请填写报告输出目录。")
      if (!state.templateItems.length) throw new Error("请至少添加一个 Excel 工作簿。")
      templates.forEach(value => validateAbsolute(value, "Word 模板", /\.(docx|docm)$/i))
      validateAbsolute(outDir, "输出目录")
      const items = state.templateItems.map((item, index) => {
        validateAbsolute(item.excel, `第 ${index + 1} 份 Excel`, /\.(xlsx|xlsm)$/i)
        if (!Number.isInteger(item.template) || item.template < 1 || item.template > templates.length) throw new Error(`第 ${index + 1} 份 Excel 的模板编号无效。`)
        return { excel: item.excel.trim(), template: item.template }
      })
      args.templates = templates; args.outDir = outDir; args.items = items
    }
    if (op.special === "blank-batch") {
      if (!state.batchItems.length) throw new Error("尚未选择要删除的空白表，请先批量扫描。")
      args.items = state.batchItems.map(item => ({ ...item, reversible: item.reversible ?? getValue("reversible", true), excel: item.excel || undefined }))
    }
    if (args.excel) validateAbsolute(args.excel, "Excel 工作簿", /\.(xlsx|xlsm)$/i)
    if (args.backupPath || args.target) {
      validateAbsolute(args.backupPath, "历史备份文件")
      validateAbsolute(args.target, "目标文件")
      if (args.backupPath.split(".").pop().toLowerCase() !== args.target.split(".").pop().toLowerCase()) throw new Error("备份文件与目标文件的扩展名必须一致。")
    }
    if (op.dryRun) args.dryRun = $("dry-run").checked
    return args
  }
  function validateAbsolute(value, label, suffix) {
    if (!isAbsolute(value)) throw new Error(`${label}必须是完整的 Windows 绝对路径。`)
    if (suffix && !suffix.test(value)) throw new Error(`${label}格式不正确。`)
  }
  function confirmationText(args) {
    const files = [args.docx, args.excel, args.target, ...(args.paths || []), ...(args.templates || [])].filter(Boolean)
    if (args.operation === "blank-batch-delete") return `将按已选清单删除 ${args.items.length} 份报告中的空白表。默认保留 Excel 表块以便恢复；Word 写盘前会备份。`
    if (args.operation === "template-batch") return `将根据 ${args.items.length} 份 Excel 生成报告，保存到“${args.outDir}”。模板原件和源工作簿不会改动。`
    if (args.operation === "backup-restore") return `将用“${args.backupPath}”覆盖“${args.target}”。恢复前会先备份目标现状。`
    return `将执行「${current().title}」，涉及：${files.length ? files.join("\n") : "当前选择的文件"}。写盘前会按选项自动备份。`
  }
  function updateRunButton() {
    const op = current()
    $("run").textContent = op.query ? op.action : $("dry-run").checked ? `预演 · ${op.action}` : op.action
    $("run").disabled = !state.initialized || !state.engine?.available || !state.engine?.versionMatched || state.busy
    document.querySelectorAll(".nav-button").forEach(button => { button.disabled = state.busy })
  }
  function setEngine(engine) {
    state.engine = engine
    const ready = Boolean(engine?.available && engine?.versionMatched)
    $("engine-pill").dataset.state = ready ? "ready" : "error"
    $("engine-text").textContent = ready ? `本地引擎 ${engine.version} 已就绪` : engine?.message || "本地引擎不可用"
    updateRunButton()
  }
  async function refreshEngine() {
    $("refresh-engine").disabled = true
    try { setEngine((await callTool("audit_engine_status")).structuredContent?.engine || { available: false, message: "未收到引擎状态。" }) }
    catch (error) { setEngine({ available: false, message: error.message }) }
    finally { $("refresh-engine").disabled = false }
  }
  function addLog(level, message) { state.logs.push({ level, message }); renderLogs() }
  function showError(message) { addLog("error", message) }
  function renderLogs() {
    $("log-empty").hidden = state.logs.length > 0
    const lines = $("log-lines"); lines.replaceChildren()
    for (const entry of state.logs.slice(-2000)) {
      const line = el("div", "log-line")
      line.dataset.level = entry.level || "info"
      line.textContent = `${entry.level === "error" ? "错误：" : entry.level === "warning" ? "警告：" : ""}${entry.message}`
      lines.appendChild(line)
    }
    $("logs").scrollTop = $("logs").scrollHeight
  }
  function renderTask() {
    const status = state.task?.status || "idle"
    const labels = { idle: "等待任务", running: "执行中", cancelling: "正在取消", succeeded: "已完成", failed: "失败", cancelled: "已取消" }
    const review = state.task?.summary?.needsReview ? " · 需复核" : ""
    $("task-status").textContent = `${labels[status] || status}${review}`
    $("task-status").dataset.status = status
    $("cancel-task").hidden = !state.task || !["running", "cancelling"].includes(status)
    if (state.task?.logs) state.logs = state.task.logs
    renderLogs()
  }
  async function execute(args) {
    if (state.busy) return
    timers.forEach(clearTimeout); timers.clear()
    state.busy = true; state.pendingAction = null; state.task = null; state.taskId = null; state.logs = []
    $("confirmation").hidden = true; $("result-card").hidden = true
    renderTask(); updateRunButton()
    try {
      const result = await callTool(current().query ? "audit_query" : "audit_run", args)
      if (result.isError || !result.structuredContent?.taskId) throw new Error(result.structuredContent?.message || result.content?.[0]?.text || "任务未能启动。")
      state.taskId = result.structuredContent.taskId
      await pollTask()
    } catch (error) {
      state.busy = false; state.task = { status: "failed", logs: state.logs }; showError(error.message); renderTask(); updateRunButton()
    }
  }
  async function pollTask() {
    if (!state.taskId) return
    timers.delete("poll")
    try {
      const result = await callTool("audit_task_status", { taskId: state.taskId })
      if (!result.structuredContent?.taskId) throw new Error(result.content?.[0]?.text || "未收到任务状态。")
      state.task = result.structuredContent; state.logs = state.task.logs || []; renderTask()
      if (["running", "cancelling"].includes(state.task.status)) {
        const timer = setTimeout(pollTask, 400); timers.set("poll", timer)
      } else {
        state.busy = false; updateRunButton()
        if (state.task.result?.code === 0) showResult(state.task.result.data)
        else if (state.task.result?.code !== undefined) showError(state.task.error || "任务未完成，请查看上方日志。")
      }
    } catch (error) {
      state.busy = false; showError(error.message); renderTask(); updateRunButton()
    }
  }
  async function cancelTask() {
    if (!state.taskId) return
    try {
      const result = await callTool("audit_cancel", { taskId: state.taskId })
      if (result.structuredContent) { state.task = result.structuredContent; renderTask() }
      if (!timers.has("poll")) { const timer = setTimeout(pollTask, 100); timers.set("poll", timer) }
    } catch (error) { showError(error.message) }
  }
  function showResult(data, title) {
    state.renderedResult = { operation: state.operation, data }
    $("result-title").textContent = title || `${current().title} · 结果`
    $("result-body").replaceChildren()
    $("result-card").hidden = false
    if (state.task?.summary) $("result-body").appendChild(renderTaskSummary(state.task.summary))
    const rendered = renderData(data, state.operation)
    if (rendered) $("result-body").appendChild(rendered)
    else $("result-body").appendChild(el("div", "notice", "操作已完成，详细结果请查看右侧日志。"))
  }
  function renderTaskSummary(summary) {
    const box = el("section", "task-summary")
    box.dataset.review = String(Boolean(summary.needsReview))
    const heading = summary.outcome === "partial" ? "部分完成，请复核" : summary.dryRun ? "预演结果" : "执行结果"
    box.appendChild(el("h4", "", heading))
    const countLabels = { tables: "表格", marks: "标注", placeholders: "占位符", tablesSynced: "同步表格", changedCells: "改动单元格", changedRows: "行结构", changedColumns: "列结构", deletedTables: "删除表格", deletedCaptions: "删除标题", restoredTables: "恢复表格", marksSynced: "同步标注", placeholdersSynced: "同步占位符", scanErrors: "扫描失败" }
    const counts = Object.entries(summary.counts || {}).filter(([, value]) => typeof value === "number")
    if (counts.length) box.appendChild(el("p", "", counts.map(([key, value]) => `${countLabels[key] || key} ${value}`).join(" · ")))
    const fileLine = (label, files) => { if (files?.length) box.appendChild(el("p", "", `${label}：${files.join("；")}`)) }
    fileLine("输出文件", summary.outputFiles)
    fileLine("涉及文件", summary.affectedFiles)
    fileLine("备份", summary.backups)
    if (summary.batch) box.appendChild(el("p", "", `批量结果：成功 ${summary.batch.succeeded}，跳过 ${summary.batch.skipped}，失败 ${summary.batch.failed}（共 ${summary.batch.total}）`))
    const reviews = [...new Set([...(summary.warnings || []), ...(summary.reviewMessages || [])])]
    if (reviews.length) {
      box.appendChild(el("p", "", `需复核 ${summary.reviewMessageCount || reviews.length} 项`))
      const list = el("ul")
      reviews.forEach(message => list.appendChild(el("li", "", message)))
      box.appendChild(list)
    }
    return box
  }
  function renderData(data, operation) {
    if (data === undefined || data === null) return null
    if (operation === "blank-scan") return renderBlankCandidates(data, "candidates")
    if (operation === "blank-batch-scan") return renderBatchCandidates(data)
    if (operation === "deleted-list") return renderDeletedList(data)
    if (operation === "table-list") return tableResult(data, { code: "编号", name: "表名", bookmark: "书签", nrows: "行数", ncols: "列数", restore_after: "原位置" })
    if (operation === "sort-list") return renderSortList(data)
    if (operation === "backup-list") return renderBackups(data)
    if (operation === "diff") return tableResult(data, { code: "表块", name: "表名", kind: "类型", row: "行", col: "列", item: "位置 / 项目", old: "Word 原值", new: "Excel 新值" }, "diff-export")
    if (operation === "template-scan") return simpleValues(data, "占位符字段")
    if (operation === "incremental-export") return el("div", "notice", "增量登记已完成；新增表格保存在「新增表格」，既有工作簿行未改动。")
    if (operation === "marks" || operation === "placeholders") return el("div", "notice", `已处理 ${data} 项。`)
    if (Array.isArray(data)) return tableResult(data, null)
    if (typeof data === "object") return tableResult([data], null)
    return el("div", "notice", String(data))
  }
  function tableResult(rows, columns, action) {
    if (!rows?.length) return el("p", "empty", "没有符合条件的记录。")
    if (!columns) {
      const keys = [...new Set(rows.flatMap(row => Object.keys(row || {})))]
      const labels = { docx: "Word", excel: "Excel", status: "状态", reason: "说明", code: "编号", name: "表名", path: "路径", added: "新增字段", timestamp: "时间", size: "大小", keepCaptionFor: "保留标题" }
      columns = Object.fromEntries(keys.slice(0, 10).map(key => [key, labels[key] || key]))
    }
    const wrap = el("div", "table-wrap")
    const table = el("table"), head = el("thead"), headRow = el("tr")
    Object.values(columns).forEach(label => headRow.appendChild(el("th", "", label)))
    head.appendChild(headRow); table.appendChild(head)
    const body = el("tbody")
    for (const row of rows) {
      const tr = el("tr")
      for (const key of Object.keys(columns)) {
        const td = el("td")
        let value = row[key]
        if (Array.isArray(value)) value = value.join("、")
        else if (value === true) value = "是"
        else if (value === false) value = "否"
        td.textContent = value === undefined || value === null || value === "" ? "—" : String(value)
        tr.appendChild(td)
      }
      body.appendChild(tr)
    }
    table.appendChild(body); wrap.appendChild(table)
    if (action === "diff-export") {
      const button = el("button", "secondary-button", "导出差异清单到 Excel"); button.type = "button"
      button.addEventListener("click", () => switchOperation("diff-export", { docx: getValue("docx"), excel: getValue("excel") }))
      const actions = el("div", "result-actions"); actions.appendChild(button); wrap.appendChild(actions)
    }
    return wrap
  }
  function simpleValues(values, label) {
    const box = el("div")
    box.appendChild(el("p", "empty", values?.length ? `共发现 ${values.length} 个${label}。` : `未发现${label}。`))
    if (values?.length) { const list = el("div", "choice-list"); values.forEach(value => list.appendChild(el("div", "choice-row", value))); box.appendChild(list) }
    return box
  }
  function renderBlankCandidates(candidates, mode = "single") {
    const box = el("div")
    if (!candidates?.length) return el("p", "empty", "未发现空白表格。")
    box.appendChild(el("p", "field-hint", "勾选要删除的表格；如需保留表标题，再勾选该行“保留标题”。"))
    const list = el("div", "choice-list")
    candidates.forEach(candidate => {
      const row = el("label", "choice-row")
      const check = el("input"); check.type = "checkbox"; check.dataset.blankNo = candidate.no
      const text = el("span", "choice-text", `${candidate.no}. ${candidate.code || "未命名表"} · ${candidate.nrows} 行 × ${candidate.ncols} 列`)
      text.appendChild(el("small", "", `${candidate.kind === "header-only" ? "仅表头有字" : "整表全空"}${candidate.caption_text ? ` · 标题：${candidate.caption_text}` : " · 未识别到标题"}`))
      const keep = el("label"); keep.style.display = "flex"; keep.style.gap = "5px"; keep.style.alignItems = "center"; keep.style.fontSize = "11px"
      const keepCheck = el("input"); keepCheck.type = "checkbox"; keepCheck.dataset.keepNo = candidate.no; keepCheck.style.width = "14px"; keepCheck.style.height = "14px"
      keep.append(keepCheck, document.createTextNode("保留标题")); row.append(check, text, keep); list.appendChild(row)
    })
    box.appendChild(list)
    const actions = el("div", "result-actions")
    const selectAll = el("button", "secondary-button", "全选"); selectAll.type = "button"
    selectAll.addEventListener("click", () => list.querySelectorAll("[data-blank-no]").forEach(input => { input.checked = true }))
    const deleteButton = el("button", "danger-button", "删除已选空白表"); deleteButton.type = "button"
    deleteButton.addEventListener("click", () => {
      const picked = [...list.querySelectorAll("[data-blank-no]:checked")].map(input => input.dataset.blankNo)
      if (!picked.length) return showError("请先勾选至少一张空白表。")
      const kept = [...list.querySelectorAll("[data-keep-no]:checked")].map(input => input.dataset.keepNo).filter(no => picked.includes(no))
      switchOperation("blank-delete", { docx: getValue("docx"), pick: picked.join(","), keepCaptionFor: kept.join(",") })
    })
    actions.append(selectAll, deleteButton); box.appendChild(actions)
    return box
  }
  function renderBatchCandidates(results) {
    const box = el("div")
    box.appendChild(el("p", "field-hint", "按报告勾选空白表；下一步会带入批量删除清单。"))
    for (const item of results) {
      const section = el("div", "form-section")
      section.appendChild(el("h3", "section-title", basename(item.docx)))
      if (item.error) { section.appendChild(el("div", "notice", `读取失败：${item.error}`)); box.appendChild(section); continue }
      if (!item.candidates?.length) { section.appendChild(el("p", "empty", "未发现空白表格。")); box.appendChild(section); continue }
      const list = el("div", "choice-list")
      for (const candidate of item.candidates) {
        const row = el("div", "choice-row")
        const check = el("input"); check.type = "checkbox"; check.dataset.docx = item.docx; check.dataset.blankNo = candidate.no
        const text = el("span", "choice-text", `候选 ${candidate.no} · ${candidate.code} · ${candidate.nrows} 行 × ${candidate.ncols} 列`)
        text.appendChild(el("small", "", candidate.caption_text || "未识别到表标题"))
        const keep = el("label"); const keepCheck = el("input"); keepCheck.type = "checkbox"; keepCheck.dataset.keepNo = candidate.no; keepCheck.title = "保留表标题"; keep.append(keepCheck, document.createTextNode("保留标题"))
        row.append(check, text, keep); list.appendChild(row)
      }
      section.appendChild(list); box.appendChild(section)
    }
    const actions = el("div", "result-actions")
    const all = el("button", "secondary-button", "全选"); all.type = "button"; all.addEventListener("click", () => box.querySelectorAll("[data-docx][data-blank-no]").forEach(input => { input.checked = true }))
    const next = el("button", "danger-button", "进入批量删除"); next.type = "button"
    next.addEventListener("click", () => {
      const grouped = new Map()
      box.querySelectorAll("[data-docx][data-blank-no]:checked").forEach(input => {
        const key = input.dataset.docx
        if (!grouped.has(key)) grouped.set(key, { docx: key, pick: [], keepCaptionFor: [] })
        const item = grouped.get(key); item.pick.push(input.dataset.blankNo)
        const keep = box.querySelector(`[data-docx="${cssEscape(key)}"] [data-keep-no="${input.dataset.blankNo}"]:checked`)
        if (keep) item.keepCaptionFor.push(input.dataset.blankNo)
      })
      if (!grouped.size) return showError("请先勾选至少一张空白表。")
      state.batchItems = [...grouped.values()].map(item => ({ ...item, pick: item.pick.join(","), keepCaptionFor: item.keepCaptionFor.join(",") }))
      switchOperation("blank-batch-delete")
    })
    actions.append(all, next); box.appendChild(actions)
    return box
  }
  function cssEscape(value) { return String(value).replace(/["\\]/g, "\\$&") }
  function renderDeletedList(items) {
    const box = el("div")
    if (!items?.length) return el("p", "empty", "工作簿中没有可恢复的已删表块。")
    const list = el("div", "choice-list")
    items.forEach((item, index) => {
      const row = el("label", "choice-row")
      const check = el("input"); check.type = "checkbox"; check.dataset.restoreNo = index + 1
      const text = el("span", "choice-text", `${index + 1}. ${item.code} ${item.name || ""}`)
      text.appendChild(el("small", "", `${item.nrows} 行 × ${item.ncols} 列 · 原位置：${item.restore_after || "末尾"}`))
      row.append(check, text); list.appendChild(row)
    })
    box.appendChild(list)
    const actions = el("div", "result-actions")
    const button = el("button", "primary-button", "恢复已选表格"); button.type = "button"
    button.addEventListener("click", () => {
      const picked = [...list.querySelectorAll("[data-restore-no]:checked")].map(input => input.dataset.restoreNo)
      if (!picked.length) return showError("请先勾选要恢复的表块。")
      switchOperation("restore-tables", { docx: getValue("docx"), excel: getValue("excel"), pick: picked.join(",") })
    })
    actions.appendChild(button); box.appendChild(actions)
    return box
  }
  function renderSortList(items) {
    const box = el("div")
    if (!items?.length) return el("p", "empty", "没有可设置的普通表块。")
    const table = tableResult(items, { code: "编号", name: "表名", sort_column: "排序列", descending: "方向", headers: "可选金额列" })
    box.appendChild(table)
    const actions = el("div", "result-actions")
    for (const item of items) {
      const edit = el("button", "secondary-button", `设置 ${item.code}`); edit.type = "button"
      edit.addEventListener("click", () => switchOperation("sort-set", { docx: getValue("docx"), excel: getValue("excel"), code: item.code, column: item.sort_column || item.headers?.[0] || "" }))
      const clear = el("button", "quiet-button", `清除 ${item.code} 排序`); clear.type = "button"
      clear.addEventListener("click", () => switchOperation("sort-clear", { docx: getValue("docx"), excel: getValue("excel"), code: item.code }))
      actions.append(edit, clear)
    }
    box.appendChild(actions)
    return box
  }
  function renderBackups(data) {
    const box = el("div")
    for (const [kind, label] of [["word", "Word 备份"], ["excel", "Excel 备份"]]) {
      box.appendChild(el("h3", "section-title", `${label}（${data[kind]?.length || 0}）`))
      if (!data[kind]?.length) { box.appendChild(el("p", "empty", "暂无历史备份。")); continue }
      const list = el("div", "choice-list")
      data[kind].forEach(item => {
        const row = el("div", "choice-row")
        const text = el("span", "choice-text", String(item.timestamp || item.time || "").replace("T", " "))
        text.appendChild(el("small", "", basename(item.path) + (item.size ? ` · ${Math.ceil(item.size / 1024)} KB` : "")))
        const restore = el("button", "secondary-button", "恢复此备份"); restore.type = "button"
        restore.addEventListener("click", () => switchOperation("backup-restore", { backupPath: item.path, target: kind === "word" ? data.word_path : data.excel_path }))
        row.append(text, restore); list.appendChild(row)
      })
      box.appendChild(list)
    }
    return box
  }
  async function scanTemplates() {
    if (state.busy) return
    const templates = listValue(getValue("templates", []))
    if (!templates.length) return showError("请先选择至少一个 Word 模板。")
    const savedTask = state.taskId
    state.busy = true; updateRunButton()
    try {
      const scans = []
      for (const docx of templates) {
        const result = await callTool("audit_query", { operation: "template-scan", docx })
        if (result.isError || !result.structuredContent?.taskId) throw new Error(result.content?.[0]?.text || "模板检查失败")
        let task
        do {
          await new Promise(resolve => setTimeout(resolve, 250))
          const status = await callTool("audit_task_status", { taskId: result.structuredContent.taskId })
          task = status.structuredContent
        } while (task && ["running", "cancelling"].includes(task.status))
        if (!task || task.result?.code !== 0) throw new Error(task?.error || `模板检查失败：${docx}`)
        state.templateScans[docx] = task.result.data || []
        scans.push({ template: basename(docx), fields: task.result.data || [] })
      }
      showResult(scans, "模板字段检查")
    } catch (error) { showError(error.message) }
    finally { state.taskId = savedTask; state.busy = false; updateRunButton() }
  }
  function onMessage(event) {
    if (event.source !== window.parent) return
    const data = event.data
    if (!data || data.jsonrpc !== "2.0") return
    if (data.id && pending.has(data.id)) { const item = pending.get(data.id); pending.delete(data.id); clearTimeout(item.timer); item.resolve(data); return }
    if (data.id && hostPending.has(data.id)) { const item = hostPending.get(data.id); hostPending.delete(data.id); clearTimeout(item.timer); item.resolve(data); return }
    if (data.method === "ui/notifications/initialized") return
    if (data.method === "ui/notifications/tool-result") {
      const task = data.params?.structuredContent
      if (task?.taskId) { state.task = task; state.taskId = task.taskId; state.logs = task.logs || state.logs; renderTask() }
      return
    }
    if (data.method === "ui/notifications/host-context-changed" && data.params?.theme) document.documentElement.dataset.theme = data.params.theme
    if (data.id !== undefined) {
      if (data.method === "ping") post({ jsonrpc: "2.0", id: data.id, result: {} })
      else if (data.method === "ui/resource-teardown") { post({ jsonrpc: "2.0", id: data.id, result: {} }); teardown() }
      else post({ jsonrpc: "2.0", id: data.id, error: { code: -32601, message: "Method not supported" } })
    }
  }
  function teardown() {
    timers.forEach(clearTimeout); timers.clear()
    for (const item of pending.values()) { clearTimeout(item.timer); item.reject(new Error("页面已关闭")) }
    pending.clear()
    for (const item of hostPending.values()) { clearTimeout(item.timer); item.reject(new Error("页面已关闭")) }
    hostPending.clear()
    window.removeEventListener("message", onMessage); window.removeEventListener("pagehide", teardown)
  }
  $("operation-form").addEventListener("submit", event => {
    event.preventDefault()
    try {
      const args = collectArgs()
      if (!current().query && !args.dryRun) {
        state.pendingAction = args; $("confirmation-title").textContent = `确认${current().title}？`; $("confirmation-text").textContent = confirmationText(args); $("confirmation").hidden = false
        $("confirmation").scrollIntoView({ block: "nearest" })
      } else execute(args)
    } catch (error) { showError(error.message) }
  })
  $("cancel-confirmation").addEventListener("click", () => { state.pendingAction = null; $("confirmation").hidden = true })
  $("confirm-run").addEventListener("click", () => { const args = state.pendingAction; if (args) execute({ ...args, confirmed: true }); state.pendingAction = null })
  $("refresh-engine").addEventListener("click", refreshEngine)
  $("cancel-task").addEventListener("click", cancelTask)
  $("clear-logs").addEventListener("click", () => { state.logs = []; renderLogs() })
  $("clear-result").addEventListener("click", () => { $("result-card").hidden = true; state.renderedResult = null })
  $("dry-run").addEventListener("change", updateRunButton)
  window.addEventListener("message", onMessage)
  window.addEventListener("pagehide", teardown)
  renderNav(); renderOperation(); renderLogs()
  if (window.parent === window) { $("engine-text").textContent = "请在 Qoder 中打开此插件页面。"; return }
  rpc("ui/initialize", { protocolVersion, appInfo: { name: "audit-toolbox", version: "__PLUGIN_VERSION__" }, appCapabilities: { availableDisplayModes: ["fullscreen"] } }, pending).then(response => {
    if (response?.error) throw new Error(response.error.message || "页面连接失败")
    if (response?.result?.protocolVersion !== protocolVersion) throw new Error("Qoder 页面协议版本不匹配，请更新客户端。")
    state.initialized = true; state.capabilities = response.result?.hostCapabilities || {}
    document.documentElement.dataset.theme = response.result?.hostContext?.theme || "light"
    post({ jsonrpc: "2.0", method: "ui/notifications/initialized", params: {} })
    updateRunButton(); refreshEngine()
  }).catch(error => { $("engine-text").textContent = error.message; $("engine-pill").dataset.state = "error" })
})()
