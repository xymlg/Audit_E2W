# 本地脚本通道

脚本只负责调用本地审计工具箱引擎并整理 JSONL 响应，不包含 Windows 引擎或运行时。使用前必须已安装审计工具箱，并允许当前 Agent 执行本地命令、访问用户选择的文件。

## 配置引擎

优先用 `--engine` 指定引擎可执行文件的绝对路径。也可设置环境变量：

```text
AUDIT_TOOLBOX_ENGINE=<审计工具箱引擎可执行文件绝对路径>
```

开发环境若用 Python 直接运行 `qoder_engine.py`，设置：

```text
AUDIT_TOOLBOX_ENGINE=<Python 可执行文件绝对路径>
AUDIT_TOOLBOX_ENGINE_ARGS=["<qoder_engine.py 绝对路径>"]
```

`AUDIT_TOOLBOX_ENGINE_ARGS` 必须是 JSON 字符串数组，不能用空格随意拆分路径。不要把本机开发路径、Python 环境或客户文件路径写进 Skill。

## 检查状态

在 Skill 根目录执行：

```bash
python scripts/audit_tool.py --engine "<引擎绝对路径>" status
```

只有返回 `ok=true`、`available=true`、`versionMatched=true` 才继续。版本不一致时先更新 Skill 或引擎，不要混用。

## 发送请求

短请求使用 `--json`：

```bash
python scripts/audit_tool.py --engine "<引擎绝对路径>" request --json '{"operation":"diff","docx":"<Word绝对路径>","excel":"<Excel绝对路径>"}'
```

长清单或批量请求先写成 UTF-8 JSON 文件，再使用 `--file`：

```bash
python scripts/audit_tool.py --engine "<引擎绝对路径>" request --file "<请求JSON绝对路径>"
```

`--json` 与 `--file` 必须且只能提供一个。请求文件只放参数，不放业务报告内容。

## 返回字段

- `ok`：引擎结果、进程退出码和 JSONL 协议是否都正常。
- `operation`、`dryRun`、`version`：本次请求信息。
- `data`：查询结果或操作返回数据。
- `logs`：引擎日志，最多保留最后 200 条；`logsTruncated=true` 表示前面还有日志。
- `protocolErrors`、`stderr`、`exitCode`：协议或进程异常。
- `error`：失败时的简要原因。

`ok=true` 仍须检查日志中的警告、跳过项和批量失败。`data` 中的整数 `0` 常表示成功码，不是改动数量。

## 执行约束

- 先发送 `dryRun=true` 或只读查询，向用户确认后再发送正式请求。
- 脚本是同步调用；命令结束前不要重复启动同一操作。
- 长任务使用当前平台的后台执行能力，并保留任务标识以便查询结果。
- 命令失败时先修复路径、权限、文件占用或版本问题，不要原样反复重试。
