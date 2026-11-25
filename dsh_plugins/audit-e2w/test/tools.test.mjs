import assert from "node:assert/strict"
import test from "node:test"
import { createTools } from "../lib/index.js"

const docx = "C:\\审计项目\\报告.docx"
const excel = "C:\\审计项目\\报告_表格数据.xlsx"
const exec = { agent: { id: "agent-1" }, callId: "call-1", signal: new AbortController().signal }

function fixture(decision = "allowed-once") {
  const calls = []
  const decisions = []
  const runtime = {
    status: async () => ({ available: true, versionMatched: true }),
    run: async request => {
      calls.push(request)
      return { result: { code: 0, data: { ok: true } }, logs: [] }
    },
  }
  const approval = {
    request: async request => {
      decisions.push(request)
      return decision
    },
  }
  return { tools: createTools({ runtime, approval }), calls, decisions }
}

test("capability catalog exposes all 28 operations", async () => {
  const { tools } = fixture()
  const result = await tools.capabilities.execute({}, exec)
  assert.equal(result.operations.length, 28)
  assert.equal(result.operations.find(item => item.operation === "writeback").dryRun, true)
})

test("read-only query does not request approval", async () => {
  const { tools, calls, decisions } = fixture()
  await tools.query.execute({ operation: "blank-scan", docx }, exec)
  assert.equal(calls[0].operation, "blank-scan")
  assert.equal(decisions.length, 0)
})

test("dry run does not request write approval", async () => {
  const { tools, calls, decisions } = fixture()
  await tools.run.execute({ operation: "export", docx, excel, dryRun: true }, exec)
  assert.equal(calls[0].dryRun, true)
  assert.equal(decisions.length, 0)
})

test("write requires allowed-once approval", async () => {
  const { tools, calls, decisions } = fixture("allowed-once")
  await tools.run.execute({ operation: "writeback", docx, excel, contents: { tables: true } }, exec)
  assert.equal(calls.length, 1)
  assert.equal(decisions.length, 1)
  assert.equal(decisions[0].agent, exec.agent)
  assert.equal(decisions[0].callId, exec.callId)
})

test("rejected write never reaches the engine", async () => {
  const { tools, calls } = fixture("rejected")
  await assert.rejects(tools.run.execute({ operation: "writeback", docx, excel }, exec), /未批准/)
  assert.equal(calls.length, 0)
})

test("relative paths and unsupported parameters fail closed", async () => {
  const { tools, calls } = fixture()
  await assert.rejects(tools.query.execute({ operation: "blank-scan", docx: "报告.docx" }, exec), /绝对路径/)
  await assert.rejects(tools.run.execute({ operation: "export", docx, shell: true }, exec), /不支持参数/)
  assert.equal(calls.length, 0)
})
