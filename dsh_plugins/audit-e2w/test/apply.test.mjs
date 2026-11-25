import assert from "node:assert/strict"
import test from "node:test"
import { apply } from "../lib/index.js"

test("plugin registers four tools and the optional skill with cleanup", async () => {
  const registered = []
  const skills = []
  const disposed = []
  const skillCtx = {
    skills: { register: skill => { skills.push(skill); return () => disposed.push("skill") } },
    effect: callback => { disposeSkill = callback() },
  }
  let disposeSkill
  const ctx = {
    tools: { register: tool => { registered.push(tool.name); return () => disposed.push(tool.name) } },
    approval: { request: async () => "rejected" },
    logger: { info() {} },
    inject: (_services, callback) => {
      callback(skillCtx)
      return { dispose: () => { disposeSkill(); disposed.push("skill-scope") } }
    },
    effect: callback => callback(),
  }
  const cleanup = apply(ctx)
  assert.deepEqual(registered, ["audit_capabilities", "audit_engine_status", "audit_query", "audit_run"])
  assert.equal(skills[0].name, "audit-word-excel")
  await cleanup()
  assert.ok(disposed.includes("audit_run"))
  assert.ok(disposed.includes("skill"))
  assert.ok(disposed.includes("skill-scope"))
})
