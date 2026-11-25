import assert from "node:assert/strict"
import test from "node:test"
import { runEngineProcess } from "../lib/engine.mjs"

const fakeEngine = `
let input = ""
process.stdin.setEncoding("utf8")
process.stdin.on("data", chunk => { input += chunk })
process.stdin.on("end", () => {
  const request = JSON.parse(input)
  console.log(JSON.stringify({ type: "log", level: "info", message: "engine-ok" }))
  console.log(JSON.stringify({ type: "result", code: 0, operation: request.operation, version: "1.11.1", data: { request } }))
})
`

test("engine JSONL bridge returns logs and result", async () => {
  const children = new Set()
  const outcome = await runEngineProcess(process.execPath, ["-e", fakeEngine], { operation: "version" }, { children })
  assert.equal(outcome.result.code, 0)
  assert.equal(outcome.result.data.request.operation, "version")
  assert.equal(outcome.logs[0].message, "engine-ok")
  assert.equal(children.size, 0)
})

test("engine JSONL bridge reassembles split output", async () => {
  const splitEngine = `
let input = ""
process.stdin.setEncoding("utf8")
process.stdin.on("data", chunk => { input += chunk })
process.stdin.on("end", () => {
  const request = JSON.parse(input)
  const line = JSON.stringify({ type: "result", code: 0, operation: request.operation, data: { ok: true } })
  process.stdout.write(line.slice(0, 20))
  setTimeout(() => process.stdout.write(line.slice(20) + "\\n"), 20)
})
`
  const outcome = await runEngineProcess(process.execPath, ["-e", splitEngine], { operation: "version" })
  assert.equal(outcome.result.data.ok, true)
})

test("engine abort rejects and clears the child", async () => {
  const children = new Set()
  const controller = new AbortController()
  const pending = runEngineProcess(process.execPath, ["-e", "setInterval(() => {}, 1000)"], { operation: "version" }, {
    signal: controller.signal,
    children,
  })
  setTimeout(() => controller.abort(), 50)
  await assert.rejects(pending, /取消/)
  assert.equal(children.size, 0)
})
