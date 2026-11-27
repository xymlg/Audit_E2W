import assert from "node:assert/strict"
import { spawn } from "node:child_process"
import { mkdir, readFile, rm, writeFile } from "node:fs/promises"
import { createInterface } from "node:readline"
import path from "node:path"
import { fileURLToPath } from "node:url"
import vm from "node:vm"

const pluginRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")
const projectRoot = path.resolve(pluginRoot, "..", "..")
const fixture = path.join(projectRoot, ".tmp", "qoder-plugin-smoke")
const backupDir = path.join(fixture, "_backup")
const backupName = "报告.20261125_120000.bak.xlsx"
const backupPath = path.join(backupDir, backupName)
await rm(fixture, { recursive: true, force: true })
await mkdir(backupDir, { recursive: true })
await writeFile(backupPath, "backup")
const manifest = JSON.parse(await readFile(path.join(pluginRoot, ".qoder-plugin", "plugin.json"), "utf8"))
const fakeEngine = path.join(fixture, "fake-engine.mjs")
await writeFile(fakeEngine, `
import readline from "node:readline"
const emit = value => process.stdout.write(JSON.stringify(value) + "\\n")
readline.createInterface({ input: process.stdin }).on("line", line => {
  const request = JSON.parse(line)
  if (request.operation === "version") {
    emit({ id: request.id, type: "result", code: 0, version: ${JSON.stringify(manifest.version)} })
    return
  }
  emit({ id: request.id, type: "log", level: "info", message: "模拟引擎开始执行" })
  if (request.operation === "export") {
    emit({ id: request.id, type: "log", level: "info", message: "发现表格：2 张" })
    emit({ id: request.id, type: "log", level: "info", message: "高亮标记：3 处" })
    emit({ id: request.id, type: "log", level: "info", message: "正文占位符：1 个" })
    emit({ id: request.id, type: "log", level: "info", message: '[导出] 已写入 报告_表格数据.xlsx（原文件已备份：${backupName}）' })
  }
  if (request.operation === "blank-batch-delete") {
    emit({ id: request.id, type: "log", level: "info", message: "批量删除空白表结束：成功 1 份，失败 0 份（共 1 份）。" })
  }
  if (request.operation === "sort-set") {
    emit({ id: request.id, type: "log", level: "error", message: "找不到列：不存在的列" })
    emit({ id: request.id, type: "result", code: 2, operation: request.operation, version: ${JSON.stringify(manifest.version)} })
    return
  }
  if (request.operation === "writeback") {
    setInterval(() => {}, 1000)
    return
  }
  setTimeout(() => {
    const data = request.operation === "blank-scan"
      ? [{ no: 1, index: 0, code: "T01", kind: "all-empty", nrows: 3, ncols: 2, caption_text: "空白表" }]
      : request.operation === "blank-batch-scan"
        ? [{ docx: request.paths[0], candidates: [{ no: 1, index: 0, code: "T01", kind: "all-empty", nrows: 3, ncols: 2, caption_text: "空白表" }], error: "" }]
        : request.operation === "diff" ? [{ code: "T01", name: "表一", kind: "单元格", row: 2, col: 1, item: "项目", old: "1", new: "2" }]
        : request.operation === "batch-writeback" ? [
            { docx: request.paths[0], excel: "D:\\smoke\\甲_表格数据.xlsx", status: "ok", reason: "同步完成" },
            { docx: "D:\\smoke\\乙.docx", excel: "D:\\smoke\\乙_表格数据.xlsx", status: "skipped", reason: "无差异" },
            { docx: "D:\\smoke\\丙.docx", excel: "D:\\smoke\\丙_表格数据.xlsx", status: "failed", reason: "工作簿不存在" },
          ]
        : request.operation === "blank-batch-delete" ? [{ docx: request.items[0].docx, status: "ok", reason: "删除完成" }]
        : undefined
    emit({ id: request.id, type: "result", code: 0, operation: request.operation, version: ${JSON.stringify(manifest.version)}, data })
  }, 20)
})
`, "utf8")

const child = spawn(process.execPath, [path.join(pluginRoot, "server.mjs")], {
  env: { ...process.env, AUDIT_TOOLBOX_ENGINE: process.execPath, AUDIT_TOOLBOX_ENGINE_ARGS: JSON.stringify([fakeEngine]) },
  stdio: ["pipe", "pipe", "pipe"],
})
const lines = createInterface({ input: child.stdout })
const pending = new Map()
let id = 0
lines.on("line", line => {
  const response = JSON.parse(line)
  const item = pending.get(response.id)
  if (item) { pending.delete(response.id); clearTimeout(item.timer); item.resolve(response) }
})
const call = (method, params = {}, timeoutMs = 5000) => new Promise((resolve, reject) => {
  const requestId = ++id
  const timer = setTimeout(() => { pending.delete(requestId); reject(new Error(`Timeout: ${method}`)) }, timeoutMs)
  pending.set(requestId, { resolve, reject, timer })
  child.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id: requestId, method, params })}\n`)
})
const flush = () => new Promise(resolve => setImmediate(resolve))
const waitForTask = async taskId => {
  let snapshot
  for (let attempt = 0; attempt < 40; attempt++) {
    snapshot = (await call("tools/call", { name: "audit_task_status", arguments: { taskId } })).result.structuredContent
    if (!["running", "cancelling"].includes(snapshot.status)) return snapshot
    await new Promise(resolve => setTimeout(resolve, 25))
  }
  throw new Error(`Task did not finish: ${snapshot?.status}`)
}

try {
  const init = await call("initialize", {
    protocolVersion: "2026-01-26",
    capabilities: { extensions: { "io.modelcontextprotocol/ui": { mimeTypes: ["text/html;profile=mcp-app"] } } },
    clientInfo: { name: "audit-toolbox-smoke", version: "1.0.0" },
  })
  assert.equal(init.result.protocolVersion, "2026-01-26")
  child.stdin.write('{"jsonrpc":"2.0","method":"notifications/initialized"}\n')
  const listed = (await call("tools/list")).result.tools
  assert.ok(listed.some(tool => tool.name === "open_audit_toolbox" && tool._meta["qoder/ui"].entrypoints.length === 2))
  assert.deepEqual(listed.find(tool => tool.name === "audit_capabilities")._meta.ui.visibility, ["model"])
  assert.deepEqual(listed.find(tool => tool.name === "audit_run")._meta.ui.visibility, ["model", "app"])
  assert.deepEqual(listed.find(tool => tool.name === "audit_query")._meta.ui.visibility, ["model", "app"])
  assert.ok(listed.find(tool => tool.name === "audit_query").inputSchema.properties.operation.enum.includes("blank-batch-scan"))
  assert.ok(listed.find(tool => tool.name === "audit_run").inputSchema.properties.operation.enum.includes("template-batch"))
  assert.ok(listed.find(tool => tool.name === "audit_query").inputSchema.properties.roleOverrides)
  assert.ok(listed.find(tool => tool.name === "audit_run").inputSchema.properties.regionStartId)
  const capabilities = (await call("tools/call", { name: "audit_capabilities", arguments: {} })).result.structuredContent
  assert.equal(capabilities.operations.length, 31)
  assert.deepEqual(capabilities.operations.find(item => item.operation === "template-batch").required, ["templates", "items", "outDir"])
  assert.equal(capabilities.operations.find(item => item.operation === "diff-export").dryRun, false)
  assert.equal(capabilities.operations.find(item => item.operation === "writeback").writes, true)
  assert.ok(capabilities.operations.find(item => item.operation === "bookmark").optional.includes("force"))
  assert.ok(!capabilities.operations.find(item => item.operation === "blank-batch-delete").optional.includes("reversible"))
  assert.equal(capabilities.operations.find(item => item.operation === "incremental-export").dryRun, true)
  assert.ok(!capabilities.operations.some(item => item.operation === "append-table"))
  assert.ok(capabilities.operations.find(item => item.operation === "format-scan").optional.includes("roleOverrides"))
  assert.ok(capabilities.operations.find(item => item.operation === "format-apply").optional.includes("regionStartId"))
  const resource = (await call("resources/read", { uri: "ui://audit-toolbox/main.html" })).result.contents[0]
  assert.equal(resource.mimeType, "text/html;profile=mcp-app")
  assert.deepEqual(resource._meta["qoder/ui"], { availableDisplayModes: ["fullscreen"], preferredDisplayMode: "fullscreen" })

  const status = (await call("tools/call", { name: "audit_engine_status", arguments: {} })).result
  assert.equal(status.isError, false, JSON.stringify(status))
  assert.equal(status.structuredContent.engine.version, manifest.version)
  const invalid = (await call("tools/call", { name: "audit_run", arguments: { operation: "export", docx: "relative.docx" } })).result
  assert.equal(invalid.isError, true)
  const unsupportedPreview = (await call("tools/call", { name: "audit_run", arguments: { operation: "diff-export", docx: "D:\\smoke\\报告.docx", dryRun: true } })).result
  assert.equal(unsupportedPreview.isError, true)
  assert.match(unsupportedPreview.content[0].text, /不支持预演/)
  const wrongParameter = (await call("tools/call", { name: "audit_run", arguments: { operation: "writeback", docx: "D:\\smoke\\报告.docx", column: "金额" } })).result
  assert.equal(wrongParameter.isError, true)
  assert.match(wrongParameter.content[0].text, /不支持参数/)
  const invalidQuery = (await call("tools/call", { name: "audit_query", arguments: { operation: "blank-delete", docx: "D:\\smoke\\报告.docx" } })).result
  assert.equal(invalidQuery.isError, true)

  const queryStarted = (await call("tools/call", { name: "audit_query", arguments: { operation: "blank-scan", docx: "D:\\smoke\\报告.docx" } })).result
  assert.equal(queryStarted.isError, false)
  const queryTask = await waitForTask(queryStarted.structuredContent.taskId)
  assert.equal(queryTask.status, "succeeded")
  assert.equal(queryTask.result.data[0].code, "T01")
  assert.ok(!Object.hasOwn(queryTask.result.data[0], "_element"))

  const started = (await call("tools/call", { name: "audit_run", arguments: { operation: "export", docx: path.join(fixture, "报告.docx"), dryRun: true } })).result
  assert.equal(started.isError, false)
  const task = await waitForTask(started.structuredContent.taskId)
  assert.equal(task.status, "succeeded")
  assert.equal(task.result.code, 0)
  assert.ok(task.logs.some(log => log.message === "模拟引擎开始执行"))
  assert.equal(task.summary.counts.tables, 2)
  assert.equal(task.summary.counts.marks, 3)
  assert.equal(task.summary.counts.placeholders, 1)
  assert.deepEqual(task.summary.outputFiles, [path.join(fixture, "报告_表格数据.xlsx")])
  assert.deepEqual(task.summary.backups, [backupPath])

  const batchStarted = (await call("tools/call", { name: "audit_run", arguments: { operation: "batch-writeback", paths: [path.join(fixture, "甲.docx")], dryRun: true } })).result
  const batchTask = await waitForTask(batchStarted.structuredContent.taskId)
  assert.equal(batchTask.status, "succeeded")
  assert.equal(batchTask.summary.outcome, "partial")
  assert.equal(batchTask.summary.batch.succeeded, 1)
  assert.equal(batchTask.summary.batch.skipped, 1)
  assert.equal(batchTask.summary.batch.failed, 1)
  assert.equal(batchTask.summary.batch.failures[0].reason, "工作簿不存在")
  assert.equal(batchTask.summary.needsReview, true)

  const blankStarted = (await call("tools/call", { name: "audit_run", arguments: { operation: "blank-batch-delete", items: [{ docx: path.join(fixture, "甲.docx"), pick: "1", keepCaptionFor: "1" }], dryRun: true } })).result
  const blankTask = await waitForTask(blankStarted.structuredContent.taskId)
  assert.equal(blankTask.summary.batch.succeeded, 1)
  assert.equal(blankTask.summary.needsReview, false)
  assert.deepEqual(blankTask.summary.affectedFiles, [path.join(fixture, "甲.docx"), path.join(fixture, "甲_表格数据.xlsx")])

  const failedStarted = (await call("tools/call", { name: "audit_run", arguments: { operation: "sort-set", docx: "D:\\smoke\\报告.docx", code: "T01", column: "金额", confirmed: true } })).result
  const failedTask = await waitForTask(failedStarted.structuredContent.taskId)
  assert.equal(failedTask.status, "failed")
  assert.equal(failedTask.summary.message, "找不到列：不存在的列")

  const longTask = (await call("tools/call", {
    name: "audit_run",
    arguments: { operation: "writeback", docx: "D:\\smoke\\报告.docx", excel: "D:\\smoke\\报告_表格数据.xlsx", contents: { tables: false, marks: true, placeholders: false, keepHighlight: true }, confirmed: true },
  })).result.structuredContent.taskId
  const cancelled = (await call("tools/call", { name: "audit_cancel", arguments: { taskId: longTask } })).result
  assert.ok(["cancelling", "cancelled"].includes(cancelled.structuredContent.status), JSON.stringify(cancelled))
  assert.equal((await waitForTask(longTask)).status, "cancelled")
  assert.equal((await call("unknown")).error.code, -32601)

  const pageScript = /<script>([\s\S]*?)<\/script>/i.exec(resource.text)?.[1]
  assert.ok(pageScript)
  const domNodes = new Map(), allNodes = []
  class FakeNode {
    constructor(selector = "") {
      this.selector = selector; this.textContent = ""; this.dataset = {}; this.style = {}; this.hidden = false; this.disabled = false; this.checked = false; this.value = ""; this.children = []; this.handlers = new Map(); this.attributes = new Map(); this.scrollTop = 0; this.scrollHeight = 0; this.className = ""; this.tag = ""
    }
    addEventListener(name, fn) { this.handlers.set(name, fn) }
    removeEventListener(name) { this.handlers.delete(name) }
    setAttribute(name, value) { this.attributes.set(name, String(value)) }
    appendChild(child) { this.children.push(child); return child }
    append(...children) { this.children.push(...children) }
    replaceChildren(...children) { this.children = children }
    dispatchEvent(event) { this.handlers.get(event.type)?.(event) }
    scrollIntoView() {}
    querySelectorAll(selector) {
      const result = []
      const walk = node => { for (const child of node.children) { if (matches(child, selector)) result.push(child); walk(child) } }
      walk(this)
      return result
    }
  }
  const matches = (node, selector) => {
    if (selector.startsWith(".")) return node.className.split(/\s+/).includes(selector.slice(1))
    const attr = /^\[([^=\]]+)(?:="([^"]*)")?\]$/.exec(selector)
    if (attr) { const actual = node.dataset[attr[1].slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase())]; return attr[2] === undefined ? actual !== undefined : String(actual) === attr[2] }
    return node.tag === selector
  }
  for (const match of resource.text.matchAll(/<([a-z]+)[^>]*\bid="([^"]+)"[^>]*>/gi)) {
    const element = new FakeNode(`#${match[2]}`)
    element.tag = match[1].toLowerCase(); element.disabled = /\bdisabled\b/i.test(match[0]); element.checked = /\bchecked\b/i.test(match[0]); element.hidden = /\bhidden\b/i.test(match[0])
    domNodes.set(element.selector, element); allNodes.push(element)
  }
  const hostEvents = new Map(), pageTimers = new Map()
  let pageTimerId = 0
  const parent = { postMessage(message) { pageSent.push(message) } }, pageSent = []
  const window = { parent, addEventListener(name, fn) { hostEvents.set(name, fn) }, removeEventListener(name) { hostEvents.delete(name) } }
  const document = {
    querySelector(selector) { return [...domNodes.values(), ...allNodes].find(node => node.selector === selector) || null },
    getElementById(idValue) { return [...domNodes.values(), ...allNodes].find(node => node.id === idValue || node.selector === `#${idValue}`) || null },
    querySelectorAll(selector) { return [...domNodes.values(), ...allNodes].filter(node => matches(node, selector)) },
    createElement(tag) { const node = new FakeNode(); node.tag = tag; allNodes.push(node); return node },
    createTextNode(text) { const node = new FakeNode(); node.textContent = text; return node },
    documentElement: { dataset: {} },
  }
  class FakeEvent { constructor(type) { this.type = type } preventDefault() {} }
  const env = {
    window, document,
    setTimeout(fn) { const key = ++pageTimerId; pageTimers.set(key, fn); return key },
    clearTimeout(key) { pageTimers.delete(key) },
    Event: FakeEvent,
    console,
  }
  vm.runInNewContext(pageScript, env)
  const handshake = pageSent.find(message => message.method === "ui/initialize")
  assert.ok(handshake)
  assert.equal(handshake.params.appInfo.version, manifest.version)
  assert.equal(handshake.params.protocolVersion, "2026-01-26")
  const deliver = data => hostEvents.get("message")({ source: parent, data: { jsonrpc: "2.0", ...data } })
  deliver({ id: handshake.id, result: { protocolVersion: "2026-01-26", hostCapabilities: { serverTools: {}, experimental: { "qoder/files": { pick: true } } }, hostContext: { theme: "dark", displayMode: "fullscreen" } } })
  await flush(); await flush()
  assert.ok(pageSent.some(message => message.method === "ui/notifications/initialized"))
  assert.equal(document.documentElement.dataset.theme, "dark")
  const engineRequest = pageSent.find(message => message.params?.name === "audit_engine_status")
  assert.ok(engineRequest)
  deliver({ id: engineRequest.id, result: { content: [{ type: "text", text: "ready" }], structuredContent: { pluginVersion: manifest.version, engine: { available: true, version: manifest.version, versionMatched: true } } } })
  await flush()
  assert.equal(domNodes.get("#run").disabled, false)
  assert.ok(domNodes.get("#task-nav").children.length >= 5)
  document.querySelectorAll(".nav-button").find(button => button.dataset.operation === "writeback").dispatchEvent(new FakeEvent("click"))
  assert.ok(document.getElementById("content-tables"), "writeback content options missing")
  assert.equal(document.getElementById("content-tables").checked, true)
  assert.equal(document.getElementById("content-keepHighlight").checked, true)
  assert.equal(document.getElementById("field-keepHighlight"), null, "writeback highlight option duplicated")
  document.querySelectorAll(".nav-button").find(button => button.dataset.operation === "export").dispatchEvent(new FakeEvent("click"))
  const docxInput = document.getElementById("field-docx")
  const excelInput = document.getElementById("field-excel")
  docxInput.value = "D:\\smoke\\报告.docx"
  docxInput.dispatchEvent(new FakeEvent("input"))
  assert.equal(excelInput.value, "D:\\smoke\\报告_表格数据.xlsx")
  domNodes.get("#operation-form").dispatchEvent(new FakeEvent("submit"))
  assert.equal(domNodes.get("#confirmation").hidden, false, domNodes.get("#log-lines").children.map(node => node.textContent).join("\n"))
  domNodes.get("#confirm-run").dispatchEvent(new FakeEvent("click"))
  await flush()
  const runRequest = pageSent.find(message => message.params?.name === "audit_run")
  assert.equal(runRequest.params.arguments.operation, "export")
  assert.equal(runRequest.params.arguments.confirmed, true)
  deliver({ id: runRequest.id, result: { content: [], structuredContent: { taskId: "page-task" } } })
  await flush()
  const taskRequest = pageSent.find(message => message.params?.name === "audit_task_status")
  deliver({ id: taskRequest.id, result: { content: [], structuredContent: { taskId: "page-task", status: "running", logs: [] } } })
  await flush()
  assert.equal(domNodes.get("#task-status").textContent, "执行中")
  assert.equal(domNodes.get("#run").disabled, true)
  assert.equal(pageTimers.size, 1)
  const [pollKey, pollTimer] = [...pageTimers.entries()][0]
  pageTimers.delete(pollKey)
  pollTimer()
  await flush()
  const secondTaskRequest = pageSent.filter(message => message.params?.name === "audit_task_status")[1]
  assert.ok(secondTaskRequest, "task polling did not continue")
  deliver({ id: secondTaskRequest.id, result: { content: [], structuredContent: { taskId: "page-task", status: "succeeded", logs: [{ level: "info", message: "页面任务完成" }], result: { code: 0, data: 0 }, summary: { outcome: "succeeded", dryRun: false, needsReview: true, counts: { changedCells: 2 }, outputFiles: ["D:\\smoke\\结果.xlsx"], affectedFiles: ["D:\\smoke\\报告.docx"], backups: ["D:\\smoke\\_backup\\报告.bak.docx"], warnings: [], reviewMessages: ["有一项需要人工复核"], reviewMessageCount: 1 } } } })
  await flush(); await flush()
  assert.equal(domNodes.get("#task-status").textContent, "已完成 · 需复核")
  assert.equal(domNodes.get("#run").disabled, false)
  assert.ok(domNodes.get("#log-lines").children.some(line => line.textContent.includes("页面任务完成")))
  const summaryNode = domNodes.get("#result-body").children[0]
  const summaryText = node => `${node.textContent || ""}${node.children.map(summaryText).join("")}`
  assert.equal(summaryNode?.className, "task-summary", `task summary missing: ${domNodes.get("#result-body").children.map(node => `${node.tag}.${node.className}`).join(" | ")}`)
  assert.match(summaryText(summaryNode), /改动单元格 2/)
  assert.match(summaryText(summaryNode), /D:\\smoke\\结果.xlsx/)
  deliver({ id: "close", method: "ui/resource-teardown" })
  await flush()
  assert.equal(pageTimers.size, 0)
  assert.equal(hostEvents.size, 0)
  assert.ok(pageSent.some(message => message.id === "close" && message.result))
  console.log("Passed: MCP protocol, all-tool registration, read-only query, engine version check, task lifecycle/cancel, page handshake, file path defaults, confirmation flow, theme, and teardown.")
} finally {
  for (const item of pending.values()) clearTimeout(item.timer)
  lines.close(); child.stdin.end(); child.kill()
  if (!process.env.KEEP_SMOKE) await rm(fixture, { recursive: true, force: true })
}
