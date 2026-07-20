# 交给其他本地 AI 使用

把公开仓库地址和本机工作簿绝对路径交给能够读取文件、运行 Python 的 AI。绝不要提供 IMAP 授权码。

可以直接使用以下提示词：

```text
请完整阅读 SKILL.md 和 references/configuration.md，并使用仓库自带脚本，为我的现有工作簿创建一套净值邮件自动化。如果我不知道 IMAP 服务器或认证方式，优先从现有客户端设置读取，再参考 references/email-providers.md 或服务商官方文档。创建运行目录并完成离线演练后，只让我在本机隐藏输入一次邮箱授权码。

Windows 使用 D:\nav-runtime 一类的短本地运行目录，不要部署在深层 OneDrive 或桌面路径。根据工作表实际结构明确选择 summary 严格汇总模式或 append 简单追加模式，不得因验证失败自动降级。summary 要求数据区后紧跟可识别汇总行；append 可初始化空白表或管理无汇总行的简单追加表，但必须配置产品名称或代码，且数据区下方不能有页脚内容。

授权码保存后运行 navctl.py propose，自动发现净值发件人、产品代码、主题规律和邮件到达时间，再读取工作簿完成路由、列映射和口径。不要让我手工抄发件人邮箱、产品代码或逐项填写 AI 能从本地证据确定的信息；只有真正歧义才问我。根据邮件历史给我一个运行安排建议，并明确询问我希望的更新频率、运行星期和本地时间，把确认结果写入 schedule；Windows 正式部署不得省略定时任务。

按表头含义识别列；summary 路由至少核对两个不同历史日期，append 冷启动要显示历史不足警告；产品、日期或数值冲突时停止。第一次生成预览给我检查，我批准并成功写入后，按我已经确认的频率和时间安装自动更新任务。后续不再要求我逐次查看预览，但每次必须先验证、备份，并在失败时保持正式表不变。夏普比率、最大回撤等分析可按一致的通行口径在未托管分析页或单独文件中完成并说明假设。
```

只克隆干净的公开仓库，不要复制已经配置过的运行目录。

## Codex（Windows）

```powershell
git clone <仓库地址> "$env:USERPROFILE\.codex\skills\nav-email-to-excel"
codex
```

然后告诉 Codex：`请使用 $nav-email-to-excel，为这个工作簿部署自动净值更新；不要让我手工整理发件人，第一次给我验收，之后自动备份写入：<绝对路径>`。

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
