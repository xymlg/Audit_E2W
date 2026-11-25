import { emptySchema, operationSpecs, pathProperties, queryOperations, runOperations, runProperties, validateRunArgs } from "../../../qoder_plugins/audit-toolbox/catalog.mjs"
import manifest from "../../../package.json" with { type: "json" }
import { AuditRuntime } from "./engine.mjs"
import { skillContent, skillDescription, skillName, skillWhenToUse } from "./skill.js"

export const name = "dsh-audit-e2w"
export const inject = ["tools", "approval"]

const output = {
  schema: { type: "object", additionalProperties: true },
  render: (_args, value) => [{ type: "text", text: JSON.stringify(value, null, 2) }],
}
const queryProperties = {
  operation: { type: "string", enum: queryOperations, description: "只读操作名；先用 audit_capabilities 查看必填参数和示例" },
  docx: pathProperties.docx,
  excel: pathProperties.excel,
  paths: pathProperties.paths,
}

export function createTools({ runtime = new AuditRuntime(), approval } = {}) {
  const capabilities = {
    name: "audit_capabilities",
    description: "只读获取 Audit_E2W 全部操作、必填参数、默认值、预演支持和调用示例。选择操作或补参数前使用。",
    parameters: emptySchema,
    output,
    execute: async () => ({
      version: manifest.version,
      operations: Object.entries(operationSpecs).map(([operation, spec]) => ({ operation, ...spec })),
    }),
  }
  const status = {
    name: "audit_engine_status",
    description: "检查本地 Audit_E2W 引擎是否可用、版本是否一致。只读，不修改文件。",
    parameters: emptySchema,
    output,
    execute: async () => runtime.status(),
  }
  const query = {
    name: "audit_query",
    description: "只读扫描 Word/Excel，不修改文件。支持空白表、可恢复表、差异、表块、排序、模板字段和备份历史。返回完整日志和结构化数据。",
    parameters: {
      type: "object",
      properties: queryProperties,
      required: ["operation"],
      additionalProperties: false,
    },
    output,
    execute: async (args, exec) => {
      const request = validateRunArgs(args, queryOperations)
      const outcome = await runtime.run(request, exec.signal)
      return { operation: request.operation, dryRun: false, ...outcome }
    },
  }
  const run = {
    name: "audit_run",
    description: "执行 Word/Excel 审计操作。先用 audit_capabilities 核对参数；写盘前先预演或差异检查，说明影响文件并取得用户确认。dryRun=true 不写文件；dryRun=false 的写操作必须通过 DSH 审批。常见映射：export=创建Excel，writeback=同步Word，template-batch=单模板报告，batch-writeback=批量同步。",
    parameters: {
      type: "object",
      properties: {
        operation: { type: "string", enum: runOperations, description: "执行操作名；先用 audit_capabilities 查看必填参数和示例" },
        ...runProperties,
      },
      required: ["operation"],
      additionalProperties: false,
    },
    output,
    execute: async (args, exec) => {
      const request = validateRunArgs(args, runOperations)
      const spec = operationSpecs[request.operation]
      if (spec.writes && request.dryRun !== true) {
        if (!exec.agent || !approval?.request) throw new Error("写操作需要 DSH 用户审批服务；当前请求已拒绝。")
        const decision = await approval.request({
          agent: exec.agent,
          toolName: "audit_run",
          callId: exec.callId,
          reason: `Audit_E2W 将执行“${spec.title}”，可能修改或生成本地 Word/Excel 文件。请核对工具参数中的文件路径。`,
          signal: exec.signal,
        })
        if (decision !== "allowed-once") throw new Error(`用户未批准写操作（${decision}）。`)
      }
      const outcome = await runtime.run(request, exec.signal)
      return { operation: request.operation, dryRun: request.dryRun === true, ...outcome }
    },
  }
  return { capabilities, status, query, run }
}

function registerSkill(ctx) {
  const fiber = ctx.inject(["skills"], skillCtx => {
    skillCtx.effect(() => skillCtx.skills.register({
      name: skillName,
      description: skillDescription,
      whenToUse: skillWhenToUse,
      content: skillContent,
      source: "bundled",
    }), "dsh-audit-e2w:skill")
  })
  return () => { void fiber.dispose() }
}

export function apply(ctx) {
  const runtime = new AuditRuntime()
  const tools = createTools({ runtime, approval: ctx.approval })
  const disposers = []
  for (const tool of Object.values(tools)) {
    disposers.push(ctx.effect(() => ctx.tools.register(tool), `dsh-audit-e2w:${tool.name}`))
  }
  disposers.push(registerSkill(ctx))
  ctx.logger.info("Audit_E2W tools mounted; local engine status is checked on demand.")
  return async () => {
    for (const dispose of disposers.reverse()) await dispose()
    runtime.dispose()
  }
}
