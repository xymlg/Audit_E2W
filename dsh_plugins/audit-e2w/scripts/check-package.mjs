import assert from "node:assert/strict"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { operationSpecs, queryOperations, runOperations } from "../../../qoder_plugins/audit-toolbox/catalog.mjs"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../..")
const manifest = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8"))
const qoderManifest = JSON.parse(readFileSync(path.join(root, "qoder_plugins/audit-toolbox/.qoder-plugin/plugin.json"), "utf8"))
const pyproject = readFileSync(path.join(root, "pyproject.toml"), "utf8")

assert.equal(manifest.name, "audit-e2w")
assert.equal(manifest.version, qoderManifest.version)
assert.match(pyproject, new RegExp(`^version = "${manifest.version.replaceAll(".", "\\.")}"$`, "m"))
assert.equal(manifest.license, "MIT")
assert.equal(manifest.repository.url, "https://github.com/xymlg/Audit_E2W.git")
assert.equal(manifest.dsh.bundle.patch, "./cordis.patch.yml")
assert.deepEqual(manifest.engines, { node: ">=24.11.0", dsh: "0.2.2-rc.5" })
assert.equal(manifest.private, undefined)
assert.ok(manifest.files.includes("dsh_plugins/audit-e2w/lib/*.mjs"), "DSH 包清单缺少 engine.mjs")
assert.equal(manifest.scripts.preinstall, undefined)
assert.equal(manifest.scripts.install, undefined)
assert.equal(manifest.scripts.postinstall, undefined)
assert.equal(queryOperations.length + runOperations.length, 28)
assert.equal(Object.keys(operationSpecs).length, 28)
assert.equal(new Set([...queryOperations, ...runOperations]).size, 28)
for (const operation of queryOperations) assert.equal(operationSpecs[operation].writes, false)
for (const operation of runOperations) assert.equal(operationSpecs[operation].writes, true)

const required = [
  "LICENSE",
  "README.md",
  "cordis.patch.yml",
  "dsh_plugins/audit-e2w/lib/index.js",
  "dsh_plugins/audit-e2w/lib/engine.mjs",
  "dsh_plugins/audit-e2w/lib/skill.js",
  "qoder_plugins/audit-toolbox/catalog.mjs",
  "excel_word/qoder_engine.py",
  "excel_word/wesync/api.py",
]
for (const item of required) assert.ok(existsSync(path.join(root, item)), `缺少 ${item}`)
const patch = readFileSync(path.join(root, "cordis.patch.yml"), "utf8")
assert.match(patch, /id: audit-e2w/)
assert.match(patch, /name: audit-e2w/)
for (const item of ["dsh_plugins/audit-e2w/lib/index.js", "dsh_plugins/audit-e2w/lib/engine.mjs", "qoder_plugins/audit-toolbox/catalog.mjs"]) {
  const source = readFileSync(path.join(root, item), "utf8")
  assert.doesNotMatch(source, new RegExp(["E:\\\\15\\.qoder", "C:\\\\Users\\\\", ["App", "Data"].join(""), "api[_-]?key\\s*[:=]\\s*[\"'][^\"']+[\"']"].join("|"), "i"))
}
console.log(`Audit_E2W ${manifest.version}: DSH package structure OK (28 operations).`)
