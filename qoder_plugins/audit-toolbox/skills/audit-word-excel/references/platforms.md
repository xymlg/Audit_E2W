# 平台安装与边界

## Qoder

优先使用审计工具箱插件提供的 MCP 工具。插件包内已包含同一份 Skill、Windows 引擎和运行时，不需要配置本地脚本通道。

如果插件已安装但工具不可用，先检查插件是否启用、版本是否一致，再重新加载插件。不要把 Qoder 的 `.qoder-plugin` 或 `.mcp.json` 复制到其他平台。

## WorkBuddy

WorkBuddy 支持个人开发者提交 Skill ZIP。ZIP 根目录应为 `audit-word-excel/`，其中包含 `SKILL.md`、`references/` 和 `scripts/`。

Skill 只提供工作方法和本地脚本，不包含审计引擎。用户本机必须另行安装审计工具箱，且 WorkBuddy 当前会话必须允许 Bash 和本地文件访问，才能真正执行文件处理。市场审核、账号认证和运行权限以 WorkBuddy 当前页面为准。

## DSH

DSH 官方支持项目级 Skill 目录：

```text
<project>/.dsh/skills/audit-word-excel/SKILL.md
```

也支持用户级 `<dshHome>/skills/audit-word-excel/SKILL.md`。将 DSH ZIP 解压到对应 skills 目录即可，不需要创建 Qoder 插件清单。

DSH 的 Skill 目录适配不等于 DSH 插件市场发布。若要通过 DSH 插件命令分发，需要另做符合其 Cordis 插件规范的包装；本包不声称已完成该包装或市场审核。

与 WorkBuddy 相同，DSH 必须允许本地命令和文件访问，并已安装审计工具箱引擎，脚本通道才可执行。
