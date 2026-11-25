import { spawn } from "node:child_process"
import { existsSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import manifest from "../../../package.json" with { type: "json" }

export const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../..")
const maxLogs = 2000

function parseLines(buffer, onLine) {
  let end
  while ((end = buffer.indexOf("\n")) !== -1) {
    const line = buffer.slice(0, end).trim()
    buffer = buffer.slice(end + 1)
    if (!line) continue
    try {
      const event = JSON.parse(line)
      if (event && typeof event === "object" && !Array.isArray(event)) onLine(event)
      else onLine({ type: "log", level: "error", message: "引擎返回了无法解析的数据。" })
    } catch {
      onLine({ type: "log", level: "error", message: `引擎返回了无法解析的数据：${line.slice(0, 200)}` })
    }
  }
  return buffer
}

export function runEngineProcess(command, args, request, { signal, cwd, timeoutMs = 600000, children } = {}) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new Error("任务已取消。"))
      return
    }
    const child = spawn(command, args, {
      cwd,
      windowsHide: true,
      shell: false,
      stdio: ["pipe", "pipe", "pipe"],
      env: { ...process.env, PYTHONUTF8: "1", PYTHONDONTWRITEBYTECODE: "1" },
    })
    children?.add(child)
    const logs = []
    let result = null
    let stderr = ""
    let buffer = ""
    let settled = false
    const addLog = (level, message) => {
      logs.push({ level, message: String(message).slice(0, 4000) })
      if (logs.length > maxLogs) logs.splice(0, logs.length - maxLogs)
    }
    const finish = (error, value) => {
      if (settled) return
      settled = true
      children?.delete(child)
      clearTimeout(timer)
      signal?.removeEventListener("abort", abort)
      if (error) reject(error)
      else resolve(value)
    }
    const abort = () => {
      child.kill()
      finish(new Error("任务已取消。"))
    }
    signal?.addEventListener("abort", abort, { once: true })
    const timer = setTimeout(() => {
      child.kill()
      finish(new Error("本地引擎执行超时（10 分钟）。"))
    }, timeoutMs)
    child.stdout.setEncoding("utf8")
    child.stderr.setEncoding("utf8")
    child.stdout.on("data", chunk => {
      buffer = parseLines(buffer + chunk, event => {
        if (event.type === "log") addLog(event.level || "info", event.message || "")
        if (event.type === "result") result = event
      })
    })
    child.stderr.on("data", chunk => {
      stderr = `${stderr}${chunk}`.slice(-16000)
      addLog("error", chunk.trim())
    })
    child.on("error", error => finish(error))
    child.on("close", code => {
      if (buffer.trim()) {
        try {
          const event = JSON.parse(buffer)
          if (event.type === "log") addLog(event.level || "info", event.message || "")
          if (event.type === "result") result = event
        } catch {
          addLog("error", "引擎结束时留下了不完整的响应。")
        }
      }
      if (settled) return
      if (!result) {
        finish(new Error(`引擎提前退出（退出码 ${code}）${stderr.trim() ? `：${stderr.trim().slice(-1000)}` : ""}`))
        return
      }
      if (result.code !== 0 || code !== 0) {
        const message = result.message || [...logs].reverse().find(item => item.level === "error")?.message || `引擎执行失败（退出码 ${code}）。`
        finish(new Error(message))
        return
      }
      finish(null, { result, logs })
    })
    child.stdin.on("error", () => {})
    child.stdin.end(`${JSON.stringify(request)}\n`, "utf8")
  })
}

function configuredExecutable(value, suffixes, label) {
  if (!value) return null
  if (!path.isAbsolute(value) || /[\u0000-\u001f]/.test(value)) throw new Error(`${label} 必须是绝对路径。`)
  if (!suffixes.some(suffix => value.toLowerCase().endsWith(suffix))) throw new Error(`${label} 文件类型不正确。`)
  if (!existsSync(value)) throw new Error(`${label} 不存在：${value}`)
  return value
}

export function resolveRuntime() {
  if (process.platform !== "win32") throw new Error("Audit_E2W 当前只支持 Windows。")
  const engine = configuredExecutable(process.env.AUDIT_TOOLBOX_ENGINE, [".exe"], "AUDIT_TOOLBOX_ENGINE")
  if (engine) return { kind: "exe", command: engine, args: [], cwd: path.dirname(engine) }
  const bundled = path.join(packageRoot, "engine", "审计工具箱引擎.exe")
  if (existsSync(bundled)) return { kind: "exe", command: bundled, args: [], cwd: path.dirname(bundled) }
  const source = path.join(packageRoot, "excel_word", "qoder_engine.py")
  if (!existsSync(source)) throw new Error("插件包缺少本地引擎源码。请重新安装 Audit_E2W。")
  const python = configuredExecutable(process.env.AUDIT_TOOLBOX_PYTHON, [".exe"], "AUDIT_TOOLBOX_PYTHON")
    || path.join(packageRoot, ".venv", "Scripts", "python.exe")
  if (existsSync(python)) return { kind: "python", command: python, args: ["-B", source], cwd: path.dirname(source) }
  return { kind: "python", command: "python", args: ["-B", source], cwd: path.dirname(source) }
}

export class AuditRuntime {
  constructor(resolver = resolveRuntime) {
    this.resolver = resolver
    this.children = new Set()
    this.cache = null
  }

  async status() {
    if (this.cache && Date.now() - this.cache.at < 5000) return this.cache.value
    let value
    try {
      const runtime = this.resolver()
      const outcome = await runEngineProcess(runtime.command, runtime.args, { id: "version", operation: "version" }, {
        cwd: runtime.cwd,
        timeoutMs: 15000,
      })
      const version = outcome.result.version
      value = {
        available: true,
        runtime: runtime.kind,
        path: runtime.command,
        version,
        versionMatched: version === manifest.version,
        message: version === manifest.version ? "引擎已就绪，版本一致。" : `插件 ${manifest.version} 与引擎 ${version} 版本不一致，请同步更新。`,
      }
    } catch (error) {
      value = { available: false, runtime: null, path: null, version: null, versionMatched: false, message: error.message }
    }
    this.cache = { at: Date.now(), value }
    return value
  }

  async run(request, signal) {
    const status = await this.status()
    if (!status.available) throw new Error(`${status.message}\n\n请安装 Audit_E2W Windows 版，或配置 AUDIT_TOOLBOX_ENGINE / AUDIT_TOOLBOX_PYTHON，并安装 requirements.txt 中的 Python 依赖。`)
    if (!status.versionMatched) throw new Error(status.message)
    const runtime = this.resolver()
    return runEngineProcess(runtime.command, runtime.args, request, { cwd: runtime.cwd, signal, children: this.children })
  }

  dispose() {
    for (const child of this.children) child.kill()
    this.children.clear()
  }
}
