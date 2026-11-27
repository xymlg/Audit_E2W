import assert from "node:assert/strict"
import test from "node:test"
import { apply } from "../lib/index.js"
import { skillContent } from "../lib/skill.js"

test("bundled skill covers format workflow and safety boundaries", () => {
  for (const required of [
    "31 个操作", "format-scan", "format-heading-candidates", "format-apply",
    "headingLevels", "roleOverrides", "regionStartId", "confirmed=true",
    "不得读取或上传整篇报告", "visual_review", "绝不覆盖原报告",
  ]) assert.match(skillContent, new RegExp(required.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")))
})

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
