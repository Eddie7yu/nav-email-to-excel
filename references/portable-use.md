# 交给其他本地 AI 使用

把公开仓库地址和本机工作簿绝对路径交给能够读取文件、运行 Python 的 AI。绝不要提供 IMAP 授权码。

可以直接使用以下提示词：

```text
请完整阅读 SKILL.md 和 references/configuration.md，并使用仓库自带脚本，为我的现有工作簿创建一套净值邮件自动化。创建运行目录后，先运行完全虚构的 navctl 离线演练；演练通过后再接入真实邮箱和工作簿配置。

Windows 使用 D:\nav-runtime 一类的短本地运行目录，不要部署在深层 OneDrive 或桌面路径。先确认每个受管数据区后紧跟可识别的汇总行；没有汇总行或存在无法安全扩展的汇总公式时停止并说明限制。

全程从只读开始。未经我明确同意，不得写入正式工作簿、发送邮件或安装定时任务。不要让我把授权码粘贴到聊天里；需要密钥时，告诉我如何亲自在本机运行隐藏输入命令。

只配置我授权的发件人和工作表。按表头含义识别列；每条路由至少核对两个不同的历史日期；产品、日期或数值存在歧义或冲突时必须停止；先生成预览和脱敏验收报告。只有我检查并批准预览后，才能通过受控的 Excel/WPS COM 流程正式写入。
```

只克隆干净的公开仓库，不要复制已经配置过的运行目录。

## Codex（Windows）

```powershell
git clone <仓库地址> "$env:USERPROFILE\.codex\skills\nav-email-to-excel"
codex
```

然后告诉 Codex：`请使用 $nav-email-to-excel，为这个工作簿部署一套预览优先的本地运行环境：<绝对路径>`。

## Claude Code（Windows）

```powershell
git clone <仓库地址> "$env:USERPROFILE\.claude\skills\nav-email-to-excel"
claude
```

然后让 Claude Code 使用 `nav-email-to-excel` Skill，并提供工作簿的本机绝对路径。

## Cursor

```powershell
git clone <仓库地址> "D:\tools\nav-email-to-excel"
cursor "D:\tools\nav-email-to-excel"
```

让 Cursor Agent 完整阅读绝对路径下的 `SKILL.md`，并且只处理另行提供的工作簿绝对路径。开始发现前，确认 Cursor 的工作区权限包含工作簿所在目录，否则它无法检查文件。

## 默认安装位置

- Codex 个人 Skill：`%USERPROFILE%\.codex\skills\nav-email-to-excel`
- Claude Code 个人 Skill：`%USERPROFILE%\.claude\skills\nav-email-to-excel`
- Cursor：任意独立的本地工具目录；明确要求 Agent 读取克隆目录中的 `SKILL.md`

更新干净安装时运行 `git -C <Skill目录> pull --ff-only`。本地运行目录应放在其他位置，避免更新 Skill 时覆盖配置。

不要把已配置的运行目录放进 AI Skill 目录或分享给别人。只分享干净的 Skill 仓库。

在 Windows PowerShell 中手动读取中文 Skill 时使用 `Get-Content -Raw -Encoding utf8 SKILL.md`，避免旧版 PowerShell 按系统代码页显示乱码。
