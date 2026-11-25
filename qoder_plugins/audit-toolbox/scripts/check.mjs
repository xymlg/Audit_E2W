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
const icon = await readFile(path.join(root, "assets/icon.png"))
assert.match(readme, /Windows 10\/11 64 位/)
assert.match(readme, /Microsoft Office/)
assert.match(readme, /不支持 macOS/)
assert.match(readme, /让 Agent 帮你干活/)
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
const { operationSpecs, queryOperations, runOperations } = catalog
assert.equal(Object.keys(operationSpecs).length, 28)
assert.match(server, /name: "audit_capabilities"/)
assert.match(server, /审计工具箱引擎\.exe/)
assert.match(html, /\/\* APP_SCRIPT \*\//)
assert.match(html, /id="task-nav"/)
assert.match(html, /id="logs"/)
assert.match(script, /ui\/initialize/)
assert.match(script, /qoder\/files\/pick/)
assert.match(script, /audit_query/)
assert.match(script, /audit_run/)
assert.equal(queryOperations.length, 8)
assert.equal(runOperations.length, 20)
for (const operation of [...queryOperations, ...runOperations]) {
  assert.ok(new RegExp(`(?:^|\\n)\\s*${operation.replace("-", "\\-")}:`).test(script) || script.includes(`"${operation}":`), `UI operation missing: ${operation}`)
}
assert.doesNotMatch(server, /https?:\/\//)
assert.doesNotMatch(script, /https?:\/\//)

const forbidden = new Set([".env", ".wesync_ui.json", "node_modules", "__pycache__", ".git"])
async function walk(directory, relative = "") {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const item = path.join(relative, entry.name)
    assert.ok(!forbidden.has(entry.name), `Forbidden package entry: ${item}`)
    assert.ok(!entry.name.startsWith("~$"), `Office lock file in package: ${item}`)
    if (entry.isDirectory()) await walk(path.join(directory, entry.name), item)
  }
}
await walk(root)
console.log(`Passed: plugin structure, versions, MCP configuration, UI markers, README/icon, and package hygiene (${manifest.version}).`)
