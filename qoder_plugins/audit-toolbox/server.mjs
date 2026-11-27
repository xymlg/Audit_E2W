import { spawn } from "node:child_process"
import { existsSync } from "node:fs"
import { access, readFile } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"

import { emptySchema, formatProperties, operationSpecs, pathProperties, queryOperations, runOperations, runProperties, validateRunArgs } from "./catalog.mjs"

const pluginRoot = path.dirname(fileURLToPath(import.meta.url))
const manifest = JSON.parse(await readFile(path.join(pluginRoot, ".qoder-plugin", "plugin.json"), "utf8"))
const pluginVersion = manifest.version
const resourceUri = "ui://audit-toolbox/main.html"
const pageScript = (await readFile(path.join(pluginRoot, "app.js"), "utf8")).replaceAll("</script", "<\\/script").replaceAll("__PLUGIN_VERSION__", pluginVersion)
const html = (await readFile(path.join(pluginRoot, "app.html"), "utf8")).replace("/* APP_SCRIPT */", () => pageScript)
const maxLogs = 2000
const tools = [
  {
    name: "open_audit_toolbox",
    title: "审计工具箱",
    description: "打开审计工具箱交互页面。用户想手动点选文件、参数并查看日志时使用；纯对话任务无需打开页面。",
    inputSchema: emptySchema,
    annotations: { readOnlyHint: true },
    _meta: {
      ui: { resourceUri, visibility: ["app"] },
      "qoder/ui": { entrypoints: [{ type: "global" }, { type: "thread" }] },
    },
  },
  {
    name: "audit_capabilities",
    title: "操作能力目录",
    description: "只读获取全部审计操作的说明、必填参数、可选参数、默认值、预演支持、写盘标志和调用示例。Agent 在选择 operation 或补参数前使用。",
    inputSchema: emptySchema,
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true },
    _meta: { ui: { visibility: ["model"] } },
  },
  {
    name: "audit_engine_status",
    title: "引擎状态",
    description: "检查本地审计工具箱引擎是否存在、版本是否与插件一致。开始任务前或出现“引擎不可用”时使用；只读，不修改文件。",
    inputSchema: emptySchema,
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ["model", "app"] } },
  },
  {
    name: "audit_query",
    title: "查询审计信息",
    description: "只读扫描，不修改文件。调用前先用 audit_capabilities 核对 operation 和参数。operation 可选 blank-scan（空白表）、blank-batch-scan（批量空白表）、diff（Excel 与 Word 差异）、table-list（表块）、sort-list（排序）、template-scan（模板占位符）、deleted-list（可恢复表）、backup-list（备份历史）、format-scan（格式预演）、format-heading-candidates（本地筛出的候选标题、表格前段落角色及附注边界，仅含有限上下文）。format-heading-candidates 不上传整篇报告或表格单元格，也不修改文件。返回 taskId 后用 audit_task_status 等待结果。",
    inputSchema: {
      type: "object",
      properties: {
        operation: { type: "string", enum: queryOperations, description: "只读操作名；先用 audit_capabilities 查看必填参数和示例" },
        ...pathProperties,
        ...formatProperties,
      },
      required: ["operation"],
      additionalProperties: false,
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true },
    _meta: { ui: { visibility: ["model", "app"] } },
  },
  {
    name: "audit_run",
    title: "执行审计操作",
    description: "执行 Word/Excel 审计操作。调用前先用 audit_capabilities 核对必填参数、默认值和预演支持。常见映射：Word 导出 Excel=export，Excel 同步回 Word=writeback，单模板生成新 Word=template-batch，批量同步=batch-writeback，格式调整=format-apply。Agent 必须先预演或差异检查，向用户说明将改动的文件，得到确认后再用 dryRun=false、confirmed=true 执行；正式写操作缺少 confirmed=true 会被拒绝。只发送 format-heading-candidates 的本地筛选结果，不上传整篇报告；标题级别、表格前段落角色和附注边界均须由用户确认，且只应用直接排版，不改标题样式、大纲级别或目录。返回 taskId 后用 audit_task_status 等待完成。",
    inputSchema: {
      type: "object",
      properties: {
        operation: { type: "string", enum: runOperations, description: "执行操作名；先用 audit_capabilities 查看必填参数、默认值和预演支持" },
        ...runProperties,
        ...formatProperties,
      },
      required: ["operation"],
      additionalProperties: false,
    },
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false },
    _meta: { ui: { visibility: ["model", "app"] } },
  },
  {
    name: "audit_task_status",
    title: "任务状态",
    description: "读取任务状态与执行日志。audit_query 或 audit_run 返回 taskId 后持续调用，直到 succeeded、failed 或 cancelled；不要只轮询一次就判断完成。",
    inputSchema: {
      type: "object",
      properties: { taskId: { type: "string", minLength: 1, maxLength: 100 } },
      required: ["taskId"],
      additionalProperties: false,
    },
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ["model", "app"] } },
  },
  {
    name: "audit_cancel",
    title: "取消任务",
    description: "取消仍在运行的本地引擎任务。已经完成的写盘不会回滚。",
    inputSchema: {
      type: "object",
      properties: { taskId: { type: "string", minLength: 1, maxLength: 100 } },
      required: ["taskId"],
      additionalProperties: false,
    },
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
    _meta: { ui: { visibility: ["model", "app"] } },
  },
]

let offersApps = false
let engineCache = { at: 0, value: null }
let activeTask = null
const tasks = new Map()

const reply = message => process.stdout.write(`${JSON.stringify({ jsonrpc: "2.0", ...message })}\n`)
const textResult = (text, structuredContent = {}, isError = false) => ({
  isError,
  content: [{ type: "text", text }],
  structuredContent,
})
const failResult = message => textResult(message, {}, true)
const isObject = value => Boolean(value) && typeof value === "object" && !Array.isArray(value)

async function exists(filePath) {
  try {
    await access(filePath)
    return true
  } catch {
    return false
  }
}

async function resolveEnginePath() {
  if (process.platform !== "win32") return null
  const candidates = []
  if (process.env.AUDIT_TOOLBOX_ENGINE) candidates.push(process.env.AUDIT_TOOLBOX_ENGINE)
  candidates.push(path.join(pluginRoot, "engine", "审计工具箱引擎.exe"))
  candidates.push(path.join(pluginRoot, "..", "..", "dist", `审计工具箱_v${pluginVersion}`, "审计工具箱引擎.exe"))
  candidates.push(path.join(process.cwd(), "dist", `审计工具箱_v${pluginVersion}`, "审计工具箱引擎.exe"))
  for (const candidate of candidates) {
    if (await exists(candidate)) return candidate
  }
  return null
}

function engineArguments() {
  if (!process.env.AUDIT_TOOLBOX_ENGINE_ARGS) return []
  const args = JSON.parse(process.env.AUDIT_TOOLBOX_ENGINE_ARGS)
  if (!Array.isArray(args) || args.some(arg => typeof arg !== "string")) throw new Error("AUDIT_TOOLBOX_ENGINE_ARGS 必须是字符串数组")
  return args
}

function parseLines(buffer, onLine) {
  let end
  while ((end = buffer.indexOf("\n")) !== -1) {
    const line = buffer.slice(0, end).trim()
    buffer = buffer.slice(end + 1)
    if (!line) continue
    try {
      const event = JSON.parse(line)
      if (isObject(event)) onLine(event)
    } catch {
      onLine({ type: "log", level: "error", message: `引擎返回了无法解析的数据：${line.slice(0, 200)}` })
    }
  }
  return buffer
}

function runEngineOnce(enginePath, request, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const child = spawn(enginePath, engineArguments(), { windowsHide: true, stdio: ["pipe", "pipe", "pipe"] })
    const events = []
    let stderr = ""
    let buffer = ""
    let settled = false
    const finish = (error, value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      if (error) reject(error)
      else resolve(value)
    }
    const timer = setTimeout(() => {
      child.kill()
      finish(new Error("引擎检查超时"))
    }, timeoutMs)
    child.stdout.setEncoding("utf8")
    child.stderr.setEncoding("utf8")
    child.stdout.on("data", chunk => {
      buffer += chunk
      buffer = parseLines(buffer, event => events.push(event))
    })
    child.stderr.on("data", chunk => { stderr += chunk })
    child.on("error", error => finish(error))
    child.on("close", code => {
      if (buffer.trim()) {
        try { events.push(JSON.parse(buffer)) } catch { /* ignore incomplete trailing data */ }
      }
      finish(null, { code, events, stderr: stderr.trim() })
    })
    child.stdin.on("error", () => {})
    child.stdin.end(`${JSON.stringify(request)}\n`, "utf8")
  })
}

async function getEngineStatus(force = false) {
  if (!force && engineCache.value && Date.now() - engineCache.at < 5000) return engineCache.value
  const enginePath = await resolveEnginePath()
  if (!enginePath) {
    const value = {
      available: false,
      pluginVersion,
      path: null,
      version: null,
      versionMatched: false,
      message: "未找到本地审计工具箱引擎，请安装最新版插件包或审计工具箱。",
    }
    engineCache = { at: Date.now(), value }
    return value
  }
  try {
    const outcome = await runEngineOnce(enginePath, { id: "version", operation: "version" })
    const result = outcome.events.find(event => event.type === "result")
    if (!result || result.code !== 0 || outcome.code !== 0) throw new Error(result?.message || outcome.stderr || `引擎无法启动（退出码 ${outcome.code}，响应 ${outcome.events.length} 条）`)
    const value = {
      available: true,
      pluginVersion,
      path: enginePath,
      version: result.version,
      versionMatched: result.version === pluginVersion,
      message: result.version === pluginVersion ? "引擎已就绪，版本一致。" : `插件 ${pluginVersion} 与引擎 ${result.version} 版本不一致，请同步更新。`,
    }
    engineCache = { at: Date.now(), value }
    return value
  } catch (error) {
    const value = {
      available: false,
      pluginVersion,
      path: enginePath,
      version: null,
      versionMatched: false,
      message: `引擎检查失败：${error.message}`,
    }
    engineCache = { at: Date.now(), value }
    return value
  }
}

function summarizeTask(task) {
  const data = task.result?.data
  const messages = task.logs.map(log => log.message)
  const unique = values => [...new Set(values.filter(value => typeof value === "string" && value))]
  const counts = {}
  const patterns = [
    ["tables", /发现表格：(\d+) 张/], ["marks", /高亮标记：(\d+) 处/],
    ["placeholders", /正文占位符：(\d+) 个/], ["bookmarksInserted", /已把 (\d+) 个书签/],
    ["markBookmarksInserted", /已为 (\d+) 处高亮标记/], ["cleanedLinks", /已删除 Word 正文中遗留的 (\d+) 处/],
    ["removedBookmarks", /已移除旧的\/锚错的工具书签 (\d+) 个/], ["clearedBookmarks", /发现本工具书签（CPAH_\/TBL_\/MRK_）：(\d+) 个/], ["rejudgedRows", /重新判定 (\d+) 行/],
    ["changedRows", /标记变化 (\d+) 行/], ["marksExported", /已刷新工作表「链接标注」：(\d+) 行标记/],
    ["marksSynced", /合计：同步 (\d+) 处标记/], ["placeholdersSynced", /合计：替换 (\d+) 处占位符/],
    ["deletedTables", /已删除 (\d+) 张空白表/], ["deletedCaptions", /已删除 \d+ 张空白表、(\d+) 个表标题/],
    ["sortedRows", /重排 (\d+) 行/], ["formulas", /合计行已写 (\d+) 个求和公式/],
    ["tablesSynced", /合计：同步 (\d+) 张表/], ["changedCells", /合计：同步 \d+ 张表（改动 (\d+) 个单元格/],
    ["changedRows", /合计：同步 \d+ 张表（改动 \d+ 个单元格，(\d+) 行结构/],
    ["changedColumns", /合计：同步 \d+ 张表（改动 \d+ 个单元格，\d+ 行结构，(\d+) 列结构）/],
    ["createdTables", /，新增 (\d+) 张表/], ["skippedTables", /，跳过 (\d+) 张表/],
    ["placeholderFieldsAdded", /本次新增 (\d+) 个/],
  ]
  for (const [key, pattern] of patterns) {
    const match = messages.map(message => pattern.exec(message)).find(Boolean)
    if (match) counts[key] = Number(match[1])
  }
  const restored = messages.filter(message => /^\[恢复\].*已重建/.test(message)).length
  if (restored) counts.restoredTables = restored
  const items = Array.isArray(data) ? data.filter(isObject) : []
  const batch = items.some(item => typeof item.status === "string")
    ? {
        total: items.length,
        succeeded: items.filter(item => item.status === "ok").length,
        skipped: items.filter(item => item.status === "skipped").length,
        failed: items.filter(item => item.status === "failed").length,
        failures: items.filter(item => item.status === "failed").map(item => ({ file: item.docx || item.excel || item.output || "", reason: item.reason || "未提供原因" })),
      }
    : null
  const defaultExcel = task.docx ? path.join(path.dirname(task.docx), `${path.parse(task.docx).name}_表格数据.xlsx`) : null
  const excel = task.excel || defaultExcel
  const outputFiles = []
  const affectedFiles = []
  if (operationSpecs[task.operation]?.writes) {
    if (task.operation === "template-batch") outputFiles.push(...items.map(item => item.output))
    else if (task.operation === "diff-export" && typeof data === "string") outputFiles.push(data)
    else if (task.operation === "backup-restore") affectedFiles.push(task.target)
    else if (task.operation === "blank-batch-delete") {
      affectedFiles.push(...(task.items || []).flatMap(item => [
        item.docx,
        item.reversible === false ? null : item.excel || path.join(path.dirname(item.docx), `${path.parse(item.docx).name}_表格数据.xlsx`),
      ]))
    } else if (["batch-writeback", "fill-placeholders"].includes(task.operation)) {
      affectedFiles.push(...items.flatMap(item => [item.docx, item.excel]))
    } else if (task.operation === "incremental-export") {
      affectedFiles.push(task.docx, excel)
    } else if (["export", "rejudge", "marks", "placeholders", "sort-set", "sort-clear"].includes(task.operation)) {
      affectedFiles.push(excel)
      if (task.operation === "export") {
        outputFiles.push(excel)
        if (messages.some(message => /^\[导出\] 已更新/.test(message))) affectedFiles.push(task.docx)
      }
      if (task.operation === "marks" && messages.some(message => /^\[标记\] 已为 .*保存 Word/.test(message))) affectedFiles.push(task.docx)
      if (["writeback", "bookmark"].includes(task.operation) && messages.some(message => /^\[定位\] 已保存/.test(message))) affectedFiles.push(excel)
    } else affectedFiles.push(task.docx)
    if (task.operation === "restore-tables") affectedFiles.push(excel)
    if (task.operation === "blank-delete" && task.reversible !== false) affectedFiles.push(excel)
  }
  const sourceFiles = unique([task.docx, excel, task.target, ...(task.items || []).flatMap(item => [item.docx, item.excel]), ...items.flatMap(item => [item.docx, item.excel, item.output])])
  const backupNames = unique(messages.flatMap(message => [
    ...Array.from(message.matchAll(/原文件已备份：([^）]+)/g), match => match[1]),
    ...Array.from(message.matchAll(/（备份：([^）]+)）/g), match => match[1]),
  ])).filter(name => name !== "无")
  const backups = backupNames.map(name => {
    const directories = unique(sourceFiles.map(file => path.dirname(file)))
    const resolved = directories.map(directory => path.join(directory, "_backup", name)).find(candidate => existsSync(candidate))
    return resolved || name
  })
  const allWarnings = unique(task.logs.filter(log => ["warn", "warning"].includes(log.level)).map(log => log.message))
  const allReviewMessages = unique([
    ...messages.filter(message => /警告|提醒|跳过|未填|保留原样|\[失败\]/.test(message)),
    ...items.filter(item => item.error).map(item => `${item.docx || item.excel || ""}：${item.error}`),
    ...(batch?.failures || []).map(item => `${item.file}：${item.reason}`),
  ])
  const warnings = allWarnings.slice(0, 20)
  const reviewMessages = allReviewMessages.slice(0, 20)
  const scanErrors = task.operation === "blank-batch-scan" ? items.filter(item => item.error).length : 0
  if (scanErrors) counts.scanErrors = scanErrors
  const failed = batch?.failed || scanErrors || (task.status === "failed" ? 1 : 0)
  return {
    operation: task.operation,
    status: task.status,
    outcome: failed ? (task.status === "succeeded" ? "partial" : task.status) : task.status,
    dryRun: task.dryRun,
    message: task.result?.message || task.error || (task.status === "succeeded" ? "任务已完成。" : ""),
    itemCount: Array.isArray(data) ? data.length : null,
    counts,
    outputFiles: unique(outputFiles),
    affectedFiles: unique(affectedFiles),
    backups,
    batch,
    warnings,
    reviewMessages,
    reviewMessageCount: allReviewMessages.length,
    needsReview: allWarnings.length > 0 || allReviewMessages.length > 0 || failed > 0,
  }
}

function taskSnapshot(task) {
  return {
    taskId: task.id,
    status: task.status,
    operation: task.operation,
    docx: task.docx,
    excel: task.excel ?? null,
    dryRun: task.dryRun,
    createdAt: task.createdAt,
    finishedAt: task.finishedAt,
    logs: task.logs,
    result: task.result,
    summary: summarizeTask(task),
    error: task.error,
  }
}

function addLog(task, level, message) {
  task.logs.push({ level, message: String(message).slice(0, 4000) })
  if (task.logs.length > maxLogs) task.logs.splice(0, task.logs.length - maxLogs)
}

function recordResult(task, event) {
  task.result = event
  task.finishedAt = new Date().toISOString()
  task.status = event.code === 0 ? "succeeded" : "failed"
  if (event.code !== 0) {
    const errorLog = [...task.logs].reverse().find(log => log.level === "error")
    task.error = event.message || errorLog?.message || "引擎返回失败，请查看执行日志。"
  }
  if (activeTask === task) activeTask = null
}

function launchTask(task, enginePath, request) {
  const child = spawn(enginePath, engineArguments(), { windowsHide: true, stdio: ["pipe", "pipe", "pipe"] })
  task.child = child
  let buffer = ""
  child.stdout.setEncoding("utf8")
  child.stderr.setEncoding("utf8")
  child.stdout.on("data", chunk => {
    buffer += chunk
    buffer = parseLines(buffer, event => {
      if (event.type === "log") addLog(task, event.level || "info", event.message || "")
      if (event.type === "result") recordResult(task, event)
    })
  })
  child.stderr.on("data", chunk => addLog(task, "error", chunk.trim()))
  child.on("error", error => {
    task.status = "failed"
    task.error = error.message
    task.finishedAt = new Date().toISOString()
    addLog(task, "error", error.message)
  })
  child.on("close", code => {
    if (buffer.trim()) {
      try {
        const event = JSON.parse(buffer)
        if (event.type === "log") addLog(task, event.level || "info", event.message || "")
        if (event.type === "result") recordResult(task, event)
      } catch { /* ignore incomplete trailing data */ }
    }
    if (task.status === "running") {
      task.status = code === 0 ? "succeeded" : "failed"
      task.finishedAt = new Date().toISOString()
      if (code !== 0 && !task.error) task.error = `引擎提前退出（退出码 ${code}）`
    }
    if (task.status === "cancelling") {
      task.status = "cancelled"
      task.finishedAt = new Date().toISOString()
    }
    task.child = null
    if (activeTask === task) activeTask = null
  })
  child.stdin.on("error", () => {})
  child.stdin.end(`${JSON.stringify({ id: task.id, ...request })}\n`, "utf8")
}

async function handle(request) {
  if (!request || request.jsonrpc !== "2.0" || typeof request.method !== "string") {
    reply({ id: null, error: { code: -32600, message: "Invalid request" } })
    return
  }
  if (request.id === undefined) return
  if (typeof request.id !== "string" && typeof request.id !== "number") {
    reply({ id: null, error: { code: -32600, message: "Invalid request identifier" } })
    return
  }
  const { id, method, params = {} } = request
  const fail = (code, message) => reply({ id, error: { code, message } })
  if (!isObject(params)) return fail(-32602, "Parameters must be an object")
  try {
    switch (method) {
      case "initialize": {
        const mimeTypes = params.capabilities?.extensions?.["io.modelcontextprotocol/ui"]?.mimeTypes
        offersApps = Array.isArray(mimeTypes) && mimeTypes.includes("text/html;profile=mcp-app")
        const protocolVersion = ["2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25", "2026-01-26"].includes(params.protocolVersion)
          ? params.protocolVersion
          : "2025-11-25"
        return reply({ id, result: {
          protocolVersion,
          capabilities: { tools: {}, resources: {} },
          serverInfo: { name: "audit-toolbox", version: pluginVersion },
        } })
      }
      case "ping": return reply({ id, result: {} })
      case "tools/list": return reply({ id, result: {
        tools: offersApps ? tools : tools.filter(tool => !tool._meta["qoder/ui"]?.entrypoints),
      } })
      case "resources/list": return reply({ id, result: {
        resources: [{ uri: resourceUri, name: "审计工具箱", mimeType: "text/html;profile=mcp-app" }],
      } })
      case "resources/read": {
        if (params.uri !== resourceUri) return fail(-32602, "Resource not found")
        return reply({ id, result: { contents: [{
          uri: resourceUri,
          mimeType: "text/html;profile=mcp-app",
          text: html,
          _meta: { "qoder/ui": { availableDisplayModes: ["fullscreen"], preferredDisplayMode: "fullscreen" } },
        }] } })
      }
      case "tools/call": {
        const args = params.arguments ?? {}
        if (!isObject(args)) return fail(-32602, "Tool arguments must be an object")
        if (params.name === "open_audit_toolbox" && Object.keys(args).length === 0) {
          const engine = await getEngineStatus()
          return reply({ id, result: textResult(
            engine.available ? `审计工具箱 ${pluginVersion}，引擎 ${engine.version} 已就绪。` : engine.message,
            { pluginVersion, engine },
          ) })
        }
        if (params.name === "audit_capabilities" && Object.keys(args).length === 0) {
          const operations = Object.entries(operationSpecs).map(([operation, spec]) => ({ operation, ...spec }))
          return reply({ id, result: textResult("审计操作目录已就绪。", {
            pluginVersion,
            operations,
            safety: { previewFirst: true, confirmBeforeWrite: true, backupByDefault: true, forceDefault: false },
          }) })
        }
        if (params.name === "audit_engine_status" && Object.keys(args).length === 0) {
          const engine = await getEngineStatus(true)
          return reply({ id, result: textResult(engine.message, { pluginVersion, engine }, !engine.available || !engine.versionMatched) })
        }
        if (params.name === "audit_query" || params.name === "audit_run") {
          const isQuery = params.name === "audit_query"
          let runRequest
          try {
            runRequest = validateRunArgs(
              args,
              isQuery ? queryOperations : runOperations,
            )
          }
          catch (error) { return reply({ id, result: failResult(error.message) }) }
          if (activeTask) return reply({ id, result: failResult("已有任务正在执行，请等待完成或先取消。") })
          const engine = await getEngineStatus(true)
          if (!engine.available || !engine.versionMatched) return reply({ id, result: failResult(engine.message) })
          if (activeTask) return reply({ id, result: failResult("已有任务正在执行，请等待完成或先取消。") })
          const taskId = `task-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
          const task = {
            id: taskId,
            status: "running",
            operation: runRequest.operation,
            docx: runRequest.docx,
            excel: runRequest.excel,
            target: runRequest.target,
            items: runRequest.items,
            reversible: runRequest.reversible,
            dryRun: runRequest.dryRun === true,
            createdAt: new Date().toISOString(),
            finishedAt: null,
            logs: [],
            result: null,
            error: null,
            child: null,
          }
          tasks.set(taskId, task)
          if (tasks.size > 20) tasks.delete(tasks.keys().next().value)
          activeTask = task
          launchTask(task, engine.path, runRequest)
          return reply({ id, result: textResult("任务已开始。", { taskId, status: task.status }) })
        }
        if (params.name === "audit_task_status" && typeof args.taskId === "string") {
          const task = tasks.get(args.taskId)
          if (!task) return reply({ id, result: failResult("找不到该任务，可能服务已重启。") })
          return reply({ id, result: textResult(`任务状态：${task.status}`, taskSnapshot(task), task.status === "failed") })
        }
        if (params.name === "audit_cancel" && typeof args.taskId === "string") {
          const task = tasks.get(args.taskId)
          if (!task) return reply({ id, result: failResult("找不到该任务，可能服务已重启。") })
          if (task.status === "running" && task.child) {
            task.status = "cancelling"
            task.child.kill()
          }
          return reply({ id, result: textResult(`任务状态：${task.status}`, taskSnapshot(task)) })
        }
        return reply({ id, result: failResult("Invalid tool name or arguments.") })
      }
      default: return fail(-32601, "Method not supported")
    }
  } catch (error) {
    reply({ id, error: { code: -32603, message: error.message || "Internal error" } })
  }
}

const maxLineBytes = 1024 * 1024
let buffer = Buffer.alloc(0)
process.stdin.on("data", chunk => {
  buffer = Buffer.concat([buffer, chunk])
  let end
  while ((end = buffer.indexOf(10)) !== -1) {
    const line = buffer.subarray(0, end)
    buffer = buffer.subarray(end + 1)
    if (line.length > maxLineBytes) {
      process.stderr.write("Request exceeds the plugin budget\n")
      process.exit(1)
    }
    if (!line.toString("utf8").trim()) continue
    try { handle(JSON.parse(line.toString("utf8"))).catch(() => reply({ id: null, error: { code: -32700, message: "Invalid JSON" } })) }
    catch { reply({ id: null, error: { code: -32700, message: "Invalid JSON" } }) }
  }
  if (buffer.length > maxLineBytes) {
    process.stderr.write("Request exceeds the plugin budget\n")
    process.exit(1)
  }
})
