# 交给其他本地 AI 使用

把公开仓库地址和本机工作簿绝对路径交给能够读取文件、运行 Python 的 AI。绝不要提供 IMAP 授权码。

可以直接使用以下提示词：

```text
请完整阅读 SKILL.md 和 references/configuration.md，并使用仓库自带脚本，为我的工作簿创建一套净值邮件自动化。已有工作簿必须原样保留；如果给定路径尚不存在，使用 `--new-workbook` 登记，完成邮箱与路由发现后运行 `workbook init-template` 创建内置完全脱敏模板，绝不覆盖已有文件。如果我不知道 IMAP 服务器或认证方式，优先从现有客户端设置读取，再参考 references/email-providers.md 或服务商官方文档。创建运行目录并完成离线演练后，只让我在本机隐藏输入一次邮箱授权码。

Windows 默认在工作簿旁创建独立的 `净值自动化` 子目录，让一份工作簿对应一套配置、入口、备份、日志和计划任务；只有默认位置过长、不可写或已有同名目录时，才用 `--destination` 另选新的短本地路径。严格遵循 references/runtime-layout.md：根目录只保留中文说明、可双击入口和 previews/backups/logs，全部程序、配置、状态、解析器和虚拟环境放进 app；不得把调试脚本、临时文件或 JSON 堆在用户根目录。自动检查工作表结构并选择合适的内部写表模式，不要让我选择或理解模式名称，也不得因验证失败擅自降低保护要求。创建运行目录后，请从 app 调用 `navctl.py secret launch` 弹出可见授权码窗口，不要只把终端命令发给我；我粘贴并回车后，请自行用 `secret status` 检查，再继续邮箱发现。

授权码保存后运行 navctl.py propose，自动发现净值发件人、产品代码、主题规律和邮件到达时间；已有表再读取工作簿完成路由、列映射和口径，新建表则从邮件规律判断日频/周频，只有明确资料证明基准时才使用有指数模板。不要让我手工抄发件人邮箱、产品代码或逐项填写 AI 能从本地证据确定的信息；只有真正歧义才问我。根据邮件历史给我一个运行安排建议，并明确询问我希望的更新频率、运行星期和本地时间，把确认结果写入 schedule；Windows 正式部署不得省略定时任务。

部署后新增产品只使用两条路径：如果我已经建立目标 Sheet，运行 `products adopt` 自动识别并接管；如果我说“格式照某个 Sheet”，运行 `products clone` 先备份，再通过 Excel/WPS 完整复制格式、清空参考产品内容并接管。不要让我填写发件人、代码、`summary/append` 或手工修改 config。新增后重新生成一次预览让我检查，原计划任务继续使用。

按表头含义识别列；只要邮箱能唯一识别产品就补入真实历史，但必须完整保留 Excel 模板。原表是周度就按周度写，不得把日度邮件全部灌入；不得清空旧历史、删除公式/累计行/基准列，或为通过验证把 summary 改成 append。若工作表是“唯一一条产品标识+预留日期、净值为空的行，下一行是累计/合计”，自动用最早真实邮件替换并补全该行，再在累计行前插入其余记录，不要问我删行或补造历史数据。完整阅读 references/template-preservation-case.md。成功运行一次 propose 后不要立即重复扫描；配置完成后直接运行 preview，让它用一次邮箱扫描完成发现、验证和预览。第一次生成预览给我检查；如果当前零新增，生成中文基线报告并在我批准后启用自动更新，不得等待下一封邮件或另写直接落表脚本。我批准预览或基线后，按确认的频率和时间安装自动更新任务。后续每次必须先验证、备份，并在失败时保持正式表不变。

不要把一次 COM 探测失败说成“电脑没装 Excel/WPS”。先检查当前是否为原生 Windows Python、读取 `doctor` 的软件探测和具体 COM 错误；必要时主动打开 Excel/WPS 完成首次启动后复查。只有虚构 `demo commit` 通过，才向我确认正式写入环境已经就绪。
```

只克隆干净的公开仓库，不要复制已经配置过的运行目录。完成前检查用户根目录没有散落的 `.py`、`.ps1`、`.json`、测试文件或诊断输出；内部排错材料必须放在 `app/diagnostics/` 并在交付前删除。

## Codex（Windows）

```powershell
git clone <仓库地址> "$env:USERPROFILE\.codex\skills\nav-email-to-excel"
codex
```

然后告诉 Codex：`请使用 $nav-email-to-excel，为这个工作簿部署自动净值更新；如果路径尚不存在就按内置脱敏模板创建。不要让我手工整理发件人，第一次给我验收，之后自动备份写入：<绝对路径>`。

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
