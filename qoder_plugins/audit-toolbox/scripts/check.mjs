import assert from "node:assert/strict"
import { access, readFile, readdir } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"

const scriptRoot = path.dirname(fileURLToPath(import.meta.url))
const root = process.argv[2] ? path.resolve(process.argv[2]) : path.resolve(scriptRoot, "..")
const skillPath = "skills/audit-word-excel/SKILL.md"
const required = [".qoder-plugin/plugin.json", ".mcp.json", "package.json", "server.mjs", "catalog.mjs", "app.html", "app.js", "README.md", "assets/icon.png", skillPath]
for (const relative of required) await access(path.join(root, relative))

const manifest = JSON.parse(await readFile(path.join(root, ".qoder-plugin/plugin.json"), "utf8"))
const packageJson = JSON.parse(await readFile(path.join(root, "package.json"), "utf8"))
const mcp = JSON.parse(await readFile(path.join(root, ".mcp.json"), "utf8"))
const readme = await readFile(path.join(root, "README.md"), "utf8")
const skill = await readFile(path.join(root, skillPath), "utf8")
const platforms = await readFile(path.join(root, "skills/audit-word-excel/references/platforms.md"), "utf8")
const icon = await readFile(path.join(root, "assets/icon.png"))
assert.match(readme, /Windows 10\/11 64 位/)
assert.match(readme, /Microsoft Office/)
assert.match(readme, /不支持 macOS/)
assert.match(readme, /让 Agent 帮你干活/)
const offPlatform = new RegExp("\\b(?:" + "Work" + "Buddy|D" + "SH)\\b|Skill\\sZIP", "i")
assert.doesNotMatch(readme, offPlatform)
assert.doesNotMatch(skill, offPlatform)
assert.match(platforms, /Qoder/)
assert.doesNotMatch(platforms, offPlatform)
assert.match(skill, /^---\r?\n[\s\S]*?\r?\n---\r?\n/)
const frontmatter = /^---\r?\n([\s\S]*?)\r?\n---\r?\n/.exec(skill)?.[1]
assert.ok(frontmatter, "Skill frontmatter missing")
for (const field of ["name", "display_name", "display_name_en", "description", "description_zh", "description_en", "version", "author", "allowed-tools"]) {
  assert.match(frontmatter, new RegExp(`^${field.replace("-", "\\-")}: .+$`, "m"), `Skill metadata missing: ${field}`)
}
assert.match(frontmatter, /^name: audit-word-excel$/m)
assert.equal(/^version:\s*(\S+)\s*$/m.exec(frontmatter)?.[1], manifest.version)
for (const field of ["description", "description_zh", "description_en"]) {
  assert.ok(new RegExp(`^${field}: (.*)$`, "m").exec(frontmatter)?.[1].length <= 500, `Skill description too long: ${field}`)
}
assert.match(skill, /scripts\/audit_tool\.py/)
assert.match(skill, /references\/operations\.md/)
assert.match(skill, /references\/local-cli\.md/)
assert.match(skill, /references\/platforms\.md/)
for (const relative of ["scripts/audit_tool.py", "references/operations.md", "references/local-cli.md", "references/platforms.md"]) {
  await access(path.join(root, "skills/audit-word-excel", relative))
}
const commandNames = ["export", "sync", "format", "batch", "check"]
const commandFiles = (await readdir(path.join(root, "commands"))).filter(file => file.endsWith(".md")).sort()
assert.deepEqual(commandFiles, commandNames.map(name => `${name}.md`).sort())
for (const name of commandNames) {
  const command = await readFile(path.join(root, "commands", `${name}.md`), "utf8")
  assert.match(command, /^---\r?\n[\s\S]*?\r?\n---\r?\n/, `Command frontmatter missing: ${name}`)
  const metadata = /^---\r?\n([\s\S]*?)\r?\n---\r?\n/.exec(command)?.[1]
  const description = /^description:\s*(.+)$/m.exec(metadata)?.[1]
  assert.ok(description && description.length <= 120, `Command description invalid: ${name}`)
  assert.match(command, /使用 `audit-word-excel` Skill/)
  for (const tool of ["audit_engine_status", "audit_capabilities", "audit_query", "audit_run", "audit_task_status"]) {
    assert.match(command, new RegExp(tool), `Command tool missing: ${name}/${tool}`)
  }
  assert.match(command, /预演|只读检查/, `Command preview missing: ${name}`)
  assert.match(command, /写盘前等待用户明确确认/, `Command approval missing: ${name}`)
  assert.match(command, /斜杠指令不是写盘确认/)
  assert.match(command, /不得自动设置 `confirmed=true`、`force=true` 或关闭备份/)
  assert.match(command, /不要用 Bash 直接调用引擎/)
  assert.match(command, /不上传整篇报告|不得读取或上传整篇报告/)
  assert.doesNotMatch(command, /https?:\/\//)
  assert.doesNotMatch(command, /audit_tool\.py|AUDIT_TOOLBOX_ENGINE|\b(?:PowerShell|cmd\.exe)\b|[A-Za-z]:\\/)
}
assert.match(skill, /audit_capabilities/)
assert.match(skill, /audit_engine_status/)
assert.match(skill, /audit_query/)
assert.match(skill, /audit_run/)
assert.match(skill, /audit_task_status/)
assert.match(skill, /先预演/)
assert.match(skill, /向用户确认/)
assert.equal(icon.subarray(0, 8).toString("hex"), "89504e470d0a1a0a")
assert.equal(icon.readUInt32BE(16), 512)
assert.equal(icon.readUInt32BE(20), 512)
assert.equal(manifest.name, "audit-toolbox")
assert.match(manifest.version, /^\d+\.\d+\.\d+$/)
assert.equal(packageJson.version, manifest.version)
assert.equal(mcp.mcpServers["audit-toolbox"].command, "${QODER_NODE_RUNTIME}")
assert.deepEqual(mcp.mcpServers["audit-toolbox"].args, ["${QODER_PLUGIN_ROOT}/server.mjs"])

const server = await readFile(path.join(root, "server.mjs"), "utf8")
const html = await readFile(path.join(root, "app.html"), "utf8")
const script = await readFile(path.join(root, "app.js"), "utf8")
assert.match(server, /ui:\/\/audit-toolbox\/main\.html/)
assert.match(server, /entrypoints: \[\{ type: "global" \}, \{ type: "thread" \}\]/)
assert.equal((server.match(/visibility: \["model", "app"\]/g) || []).length, 5)
const catalog = await import(pathToFileURL(path.join(root, "catalog.mjs")).href)
const { operationSpecs, queryOperations, runOperations, validateRunArgs } = catalog
const sampleDocx = path.join(root, "报告.docx")
const formatArgs = { operation: "format-apply", docx: sampleDocx, dryRun: true }
assert.equal(validateRunArgs(formatArgs, runOperations).dryRun, true)
assert.throws(() => validateRunArgs({ ...formatArgs, dryRun: false }, runOperations), /confirmed=true/)
assert.equal(validateRunArgs({ ...formatArgs, dryRun: false, confirmed: true }, runOperations).confirmed, true)
for (const operation of runOperations) {
  assert.ok(operationSpecs[operation].optional.includes("confirmed"), `Write confirmation missing: ${operation}`)
  const { confirmed: _confirmed, ...unconfirmedExample } = operationSpecs[operation].example
  const writeArgs = { ...unconfirmedExample, ...(operationSpecs[operation].dryRun ? { dryRun: false } : {}) }
  assert.throws(() => validateRunArgs(writeArgs, runOperations), /confirmed=true/, operation)
}
assert.ok(runOperations.includes("incremental-export"))
assert.ok(!runOperations.includes("append-table"))
assert.throws(() => validateRunArgs({ operation: "format-scan", docx: sampleDocx, scope: { headings: true } }, queryOperations), /七项/)
assert.throws(() => validateRunArgs({ operation: "format-scan", docx: sampleDocx, headingLevels: { p0003: 1 } }, queryOperations), /同时提供/)
assert.throws(() => validateRunArgs({ operation: "format-scan", docx: sampleDocx, headingBatchId: "batch", headingLevels: { p0003: 5 } }, queryOperations), /1～4/)
assert.throws(() => validateRunArgs({ operation: "format-scan", docx: sampleDocx, roleOverrides: { b0008: "body" } }, queryOperations), /同时提供/)
assert.throws(() => validateRunArgs({ operation: "format-scan", docx: sampleDocx, roleBatchId: "batch", roleOverrides: { b0008: "heading_1" } }, queryOperations), /body\/table_title/)
assert.throws(() => validateRunArgs({ operation: "format-scan", docx: sampleDocx, regionStartId: "b0000" }, queryOperations), /roleBatchId/)
const confirmedRole = validateRunArgs({ operation: "format-scan", docx: sampleDocx, roleBatchId: "batch", roleOverrides: { b0008: "table_title" }, regionStartId: "b0000" }, queryOperations)
assert.deepEqual(confirmedRole.roleOverrides, { b0008: "table_title" })
assert.equal(confirmedRole.regionStartId, "b0000")
assert.equal(Object.keys(operationSpecs).length, 31)
assert.match(server, /name: "audit_capabilities"/)
assert.match(server, /审计工具箱引擎\.exe/)
assert.match(html, /\/\* APP_SCRIPT \*\//)
assert.match(html, /id="task-nav"/)
assert.match(html, /id="logs"/)
assert.match(script, /ui\/initialize/)
assert.match(script, /qoder\/files\/pick/)
assert.match(script, /audit_query/)
assert.match(script, /audit_run/)
assert.equal(queryOperations.length, 10)
assert.equal(runOperations.length, 21)
for (const operation of [...queryOperations, ...runOperations]) {
  if (operationSpecs[operation].agentOnly) continue
  assert.ok(new RegExp(`(?:^|\\n)\\s*${operation.replace("-", "\\-")}:`).test(script) || script.includes(`"${operation}":`), `UI operation missing: ${operation}`)
}
assert.match(script, /keepCaptionFor: item\.keepCaptionFor\.join\(","\)/)
assert.match(script, /execute\(\{ \.\.\.args, confirmed: true \}\)/)
assert.doesNotMatch(server, /https?:\/\//)
assert.doesNotMatch(script, /https?:\/\//)

const forbidden = new Set([".env", ".wesync_ui.json", "node_modules", "__pycache__", ".git"])
const textSuffixes = new Set([".md", ".txt", ".json", ".mjs", ".js", ".html", ".py"])
async function walk(directory, relative = "") {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const item = path.join(relative, entry.name)
    assert.ok(!forbidden.has(entry.name), `Forbidden package entry: ${item}`)
    assert.ok(!entry.name.startsWith("~$"), `Office lock file in package: ${item}`)
    if (entry.isDirectory()) {
      await walk(path.join(directory, entry.name), item)
    } else if (textSuffixes.has(path.extname(entry.name).toLowerCase())) {
      const text = await readFile(path.join(directory, entry.name), "utf8")
      assert.doesNotMatch(text, offPlatform, `Off-platform publishing content: ${item}`)
    }
  }
}
await walk(root)
console.log(`Passed: plugin structure, versions, MCP configuration, ${commandNames.length} slash commands, UI markers, README/icon, and package hygiene (${manifest.version}).`)
