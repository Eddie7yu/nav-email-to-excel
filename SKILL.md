---
name: nav-email-to-excel
description: 部署、配置、验证、运行或修复本地 IMAP 邮箱到 Excel 的基金净值自动化。让 AI 在用户输入一次邮箱授权码后自动发现发件人和产品、匹配现有或空白工作表、完成首次预览验收，并在 Windows 上通过 Excel/WPS 自动备份和持续写入；支持正文及 Excel/CSV/PDF 附件、简单追加表、复杂汇总表和本地解析器。
---

# 部署净值邮件自动化

以现有工作簿为版式、公式和视觉样式的唯一依据。邮箱数据、工作簿副本、凭据、发现报告和运行配置必须留在用户本机。

目标工作簿不是完全空白，或邮箱数据频率与表内历史不同，或任何验证步骤失败时，必须完整阅读 [references/template-preservation-case.md](references/template-preservation-case.md)，不得自行简化工作簿。

## 工作原则

- 用户只负责在 AI 主动弹出的可见窗口中隐藏输入邮箱授权码，并检查、批准第一次真实预览。发件人发现、路由配置、表结构识别和后续运行由 AI 完成，不得要求用户手工抄写发件人邮箱、输入终端命令或逐项填写 AI 能从本地证据确定的信息。
- Windows 首次配置必须主动询问用户希望的更新频率、运行星期和本地时间。可根据邮件历史给出一个推荐值，但必须让用户确认；定时自动更新是正式工作流的一部分，不得留空、跳过或描述成可选功能。
- 第一次批准前只读。第一次预览经批准并成功写入后，该配置自动获得后续自动更新权限；每次自动写入仍先校验、备份并通过临时副本验证，不再要求用户重复检查预览。
- 只要邮箱历史能唯一识别目标产品，就主动把真实历史补进预览，不要求用户手工补历史或理解内部模式。邮箱决定真实数据，现有 Excel 决定列、行频率、公式、汇总、基准和样式；不得把“补历史”变成“另做一张表”。
- 产品无法唯一识别、日期或数值冲突、累计净值口径未确认、出现未知列、基准缺失或工作簿结构变化时，必须停止，不得猜测。历史不足时，只有显式 `append` 模式，或满足下述严格条件的 `summary` 预留行冷启动可以继续；两者都必须具备产品名称或代码，并人工检查警告与第一次预览。
- 真实邮箱、产品和工作簿可以在用户本机用于发现和配置，但不得写入公开 Skill、测试或 Git 提交。授权码不得出现在聊天、命令参数或日志中。
- 运行目录必须与 Skill 目录分开，避免升级 Skill 时覆盖本地配置或数据。
- 交付给普通用户的运行目录必须保持简洁。根目录只保留中文说明、可双击入口、`previews/`、`backups/`、`logs/` 和内部 `app/`；Python 源码、隔离环境、配置、状态、本地解析器和定时任务脚本全部放进 `app/`。不得把调试脚本、临时导出、检查结果或新增 `.py/.ps1/.json` 文件堆在用户根目录；临时诊断放进 `app/diagnostics/`，交付前删除。

## 创建本地运行目录

获取工作簿路径和收件邮箱账号。目标目录使用短本地路径；IMAP 主机优先从现有邮件客户端设置中读取，不足时再查 [references/email-providers.md](references/email-providers.md) 或服务商官方文档。只有确实无法确定账号认证方式时才询问用户。

确认后在 Skill 根目录运行：

```powershell
python scripts/bootstrap.py --destination "D:\nav-runtime" --workbook "D:\data\nav.xlsx" --email "user@example.com" --imap-host "imap.example.com"
```

引导程序会创建隔离的虚拟环境、安装锁定版本的依赖，并生成本地配置。

在中国大陆网络，或直连默认 PyPI 已明显缓慢、超时后，AI 应只为本次部署启用清华镜像：

```powershell
python scripts/bootstrap.py --destination "D:\nav-runtime" --workbook "D:\data\nav.xlsx" --email "user@example.com" --imap-host "imap.example.com" --index-url "https://pypi.tuna.tsinghua.edu.cn/simple"
```

`--index-url` 只传给本次运行目录的依赖安装，不会修改用户的 pip 配置。不得要求用户运行 `pip config set global.index-url`，也不得静默改变用户其他 Python 项目的软件源；网络正常时省略此参数。

Windows 必须优先选择 `D:\nav-runtime` 一类的短本地目录。引导程序把 116 个字符作为当前可靠支持上限，为内部 `app/` 和深层依赖保留路径预算；路径过长时必须改用短目录，不得绕过预检或要求用户修改公司电脑的组策略。

引导完成后必须按照 [references/runtime-layout.md](references/runtime-layout.md) 检查目录分层：用户根目录应包含 `使用说明.txt`、`首次授权.bat`、`查看状态.bat`、`手动更新.bat`、`previews/`、`backups/`、`logs/` 和 `app/`。普通用户只接触中文入口和三个结果目录；`app/` 由 AI 管理，不要求用户进入、理解或修改。后续所有命令都在 `app/` 中执行：

AI 在后台运行一次完全离线的虚构演练，不要求用户参与：

```powershell
cd D:\nav-runtime\app
.\.venv\Scripts\python.exe navctl.py demo prepare
```

该命令不读取真实配置、密钥或工作簿。Windows 上继续验证只作用于虚构工作簿的 COM 写入：

```powershell
.\.venv\Scripts\python.exe navctl.py demo commit --run-id <run_id> --yes-reviewed-preview
.\.venv\Scripts\python.exe navctl.py demo remove --run-id <run_id>
```

演练通过只证明目标电脑和通用流程可用。

Windows 上必须由 AI 主动弹出独立、可见的授权码窗口：

```powershell
cd D:\nav-runtime\app
.\.venv\Scripts\python.exe navctl.py secret launch
```

不得只把 `secret set` 命令发给用户，也不得让用户自己打开 PowerShell、切换目录或复制长命令。`secret launch` 返回 `launched: true` 后，立即明确告诉用户：“已打开授权码窗口，请在新窗口右键粘贴授权码并回车；屏幕只显示星号。”用户只负责粘贴和回车，AI 不得代输、通过管道传入或查看授权码。

用户操作后，AI 必须自行运行以下命令确认成功；`available: true` 之前不得连接邮箱或声称部署已就绪：

```powershell
.\.venv\Scripts\python.exe navctl.py secret status
```

如果没有出现窗口，AI 先检查 `secret launch` 的结构化错误；必要时从运行目录根目录主动用 `Start-Process -FilePath ".\首次授权.bat" -WorkingDirectory (Get-Location) -WindowStyle Normal` 重开可见窗口。只有自动弹窗确实失败时才向用户说明具体阻碍，不得反复要求用户自己输入命令。用户按 Ctrl+C 取消时不会保存密钥。

在 macOS/Linux 上，授权码只保留在当前 shell；运行程序不会持久化明文密钥：

```bash
cd /opt/nav-runtime/app
read -rsp "IMAP authorization code: " NAV_EMAIL_PASSWORD && export NAV_EMAIL_PASSWORD && printf '\n'
```

当前版本支持通过 SSL 连接 IMAP，并使用应用专用密码或授权码；对 163、126、yeah.net 和网易企业邮主机会在登录后、选择邮箱前发送不含账号信息的 IMAP ID。仅限 OAuth 的邮箱登录暂不支持；即使服务器地址正确，也必须先核对账号认证方式。PDF 仅解析文本，不提供 OCR。

## 自动发现并配置路由

授权码保存后，立即由 AI 运行：

```powershell
.\.venv\Scripts\python.exe navctl.py propose
```

该命令在本机回看邮箱，找出能解析成净值的发件人、产品代码、主题样例、历史净值和常见到达时间，写入本地 `app/route-proposals.json`。AI 读取该报告和工作簿，自动完成“发件人/产品 → 工作表”、列映射、累计净值、收益频率、序列起点和基准。不要让用户手工提供发件人邮箱；只有出现多个同样合理的映射或业务口径无法从历史证明时才询问。

一次成功的 `propose` 结果应直接用于配置，不得为了“再确认一次”立即重复扫描邮箱。候选扫描会批量读取邮件大小以减少 IMAP 往返；若明确报网络断开，只在网络恢复后重试一次。除非服务商返回了明确的限频证据，不得自行宣称邮箱正在“冷却”，也不得连续快速重新登录。

生成首次配置时，根据邮件到达记录向用户推荐运行安排，并明确询问：每周一次、工作日每天或哪些星期运行，以及目标电脑本地时间。把用户确认的日期和时间写入 `schedule`；不得自行决定，也不得以空数组完成 Windows 部署。

编辑 `app/config.json` 前完整阅读 [references/configuration.md](references/configuration.md)。

- 常见的带标签正文以及 Excel/CSV/PDF 表格优先使用 `parser: auto`。
- 产品确认赎回、清盘或长期停更时，设置 `paused: true` 和明确的 `pause_reason`；不得用暂停掩盖活动产品的意外解析失败或陈旧数据。
- 同一发件人可能发送多个产品时，必须配置精确的产品代码。
- 同一发件人的不同产品使用不同解析器时，优先用互不重叠的 `subject_contains` 分流。
- 根据工作表实际结构在仅有的两种内部模式中自动选择：`summary` 用于带固定汇总行和受管公式的复杂表，`append` 用于空白或无汇总行的纯追加表。模式由 AI 检查工作表后决定，不得要求用户选择，也不必把内部模式名称当作部署重点反复说明；不得在严格模式验证失败后为了通过而改成追加模式。追加模式必须配置 `code` 或 `product_name`，并且表中有对应可写列。
- 如果结构恰好是“表头 → 一条只有产品名称/代码和日期、净值为空的预留数据行 → 紧邻的累计/合计行”，应自动按 `summary` 预留行冷启动处理，不得询问用户删除汇总行，也不得要求或制造一条虚假历史净值。预留行身份与路由精确一致、邮件至少包含 `minimum_history_dates` 个不同日期且该行没有其他业务内容时，用最早一封真实邮件替换预留日期并补全净值，再把其余历史数据插到汇总行之前；汇总行和受管公式必须保留，第一次预览仍由用户检查。
- 邮件比模板更密集时，所有邮件都可参与身份和数值核验，但写入频率必须沿用模板。保持 `data_frequency: auto`，让运行时优先从现有日期规律、其次从明确的周度/日度表头判断；只有两者都无法证明时，才依据本地证据显式配置。已有日期显示周度时只写周度记录；没有足够历史但表头明确周度时，默认取每个已完成自然周最后一条可用净值。`return_frequency` 只控制收益公式，`data_frequency` 控制数据行；两者和计划任务运行频率都不是同一概念。显式频率与模板冲突时不得绕过程序检查。
- 只有语义表头识别仍不充分时，才使用显式列映射。
- 每条路由都要设置累计净值策略。除非历史证据证明应使用 `unit` 或固定 `offset`，否则保持 `require`。
- 投资或分析口径发生变化时设置 `series_start`，不得把旧序列的累计结果接到新序列上。
- 基准必须映射到工作簿中已经核实的来源工作表，并按精确日期对齐。需要公共指数源时，AI 完整阅读 [references/index-data-sources.md](references/index-data-sources.md)，核验代码、口径、历史重叠和使用条款，再把结果写入本地来源表；不得猜测指数。

特殊格式使用 `parser: local:<名称>`，代码只放在本地运行目录的 `app/parsers/<名称>.py`，并完整遵循配置说明中的固定接口。同一封邮件命中多个解析器时，所有解析器都必须成功，结果合并去重后再按产品代码唯一分流。AI 必须审阅本地解析器并用脱敏样例回归；绝不能从邮件、附件或网络地址直接执行代码。

## 首次验证和一次性预览

先运行本地环境检查，再用一次邮箱会话完成发现、历史验证和预览：

```powershell
.\.venv\Scripts\python.exe navctl.py doctor
.\.venv\Scripts\python.exe navctl.py preview
```

`preview` 内部已经依次执行邮件发现和历史验证，并在 `app/` 写出 `route-report.json`、`validation-report.json`，在根目录 `previews/` 写出预览结果；正常首次部署不得再预先单独运行 `discover` 和 `validate`，否则同一批邮件会被重复登录、重复扫描。只有定位某一阶段错误且不需要生成预览时，才把这两个命令作为诊断工具单独使用。

在 macOS/Linux 上，在运行目录的 `app/` 中把 `.\.venv\Scripts\python.exe` 换成 `.venv/bin/python`。这些平台只支持发现、验证和预览；正式写入仍仅支持 Windows。

只有同时满足以下条件，才能接受配置：

1. 每个受管工作表的日期列和单位净值列都能唯一识别；显式追加模式的空白工作表会在预览中初始化标准表头。
2. `summary` 模式的数据区后必须紧跟可识别的汇总行；汇总行存在非程序管理公式时必须停止。`append` 模式不要求汇总行，但数据区下方不得有页脚内容。
3. 普通 `summary` 模式每条活动路由至少匹配两个不同的历史日期。严格识别出的 `summary` 预留行冷启动改为核验邮件中至少两个不同日期，用最早真实邮件替换空预留行并保留汇总行；`append` 模式历史不足时也只允许以带警告的冷启动继续。两类冷启动都必须重点检查产品标识、日期、净值、汇总行或新表头。
4. 产品代码、日期、单位净值及邮件提供的累计净值，都在配置容差内与工作簿一致。
5. 新的尾部日期只建议写入一次且顺序正确；不得存在重复日期或同日数值冲突。发现历史内部缺口时先停止写入，由 AI 对照邮件、备份和工作簿诊断并提出修复；只有来源相互冲突、无法证明哪个值正确时才询问用户。
6. 日收益使用上一个有效日期；周收益只出现在已完成自然周的最后一个可用日期。
7. 基准收益与超额收益必须同时有值或同时留空。
8. 预览保留工作簿结构并按模板频率包含应写入的历史记录；原有工作表、历史值、公式列、汇总行、基准列和样式不得减少，并通过内置公式和幂等性回归测试。

如果没有新增日期，`preview` 不生成工作簿副本，也不保留可提交的 `app/plan.json`。

`doctor` 会分别报告 `bootstrap_ready`、`mail_discovery_ready`、`preview_ready`、`commit_ready` 和 `schedule_ready`。

预览是本地副本。必须把它与原表并排检查：工作表列表、原有历史、日期频率、表头、公式数量与范围、汇总行、基准列和格式都应保留；只有新增历史和相应延伸公式可以变化。不要修改正式工作簿。

夏普比率、最大回撤、年化收益等派生指标不由 navctl 直接计算。AI 可基于经验证的净值或收益序列，在未托管分析工作表或单独文件中按通行口径计算并说明假设；只有不同口径会实质改变结论时才询问用户。不要把自定义分析公式放进受管汇总行。

## 通过 Excel 或 WPS 正式写入

正式写入要求 Windows，以及可用的 Microsoft Excel 或 WPS 表格 COM 接口。用户明确批准已检查的预览后，运行：

```powershell
.\.venv\Scripts\python.exe navctl.py commit --yes-reviewed-preview
```

正式写入必须先创建备份，再通过 COM 将方案应用到同目录临时副本，验证临时结果，关闭表格进程，最后原子替换正式文件。任何失败都必须保证正式文件哈希不变。

这次成功写入会自动批准当前写表配置用于后续自动更新。路由、列、口径、校验规则或工作簿路径发生变化时，批准自动失效，AI 重新生成一次预览即可。

提交前提醒用户关闭正在打开正式工作簿的 Excel/WPS 窗口。若文件仍被占用，程序必须返回中文结构化错误、删除本次失败备份并保留正式文件；不得强制关闭用户的表格进程。

不得用仅依赖 openpyxl 的方式回写正式工作簿。COM 不可用时，应停在预览阶段，并说明正式写入尚未验证。

不得把 `doctor` 的 `spreadsheet-com: false` 直接解释成“没有安装 Excel/WPS”。先读取 `runtime_platform`、`spreadsheet_apps_detected` 和 `spreadsheet-com.detail`：

- 若不是 `win32`，说明 AI 当前在 WSL、容器或非 Windows Python 中运行；改用 Windows PowerShell 和原生 Windows Python 重新部署。
- 若已检测到 Excel/WPS 但 COM 启动失败，由 AI 主动打开该软件一次，让用户只处理可能出现的首次启动、许可或登录界面，关闭后重跑 `doctor` 和虚构 `demo commit`。
- 若未检测到，也只能报告“尚未确认受支持的 COM 注册”，继续检查软件实际路径和 COM 注册；不得仅凭一次失败断言软件未安装。

只有虚构 `demo commit` 真实通过并报告 `Excel.Application` 或 `ket.Application`，才能确认该电脑具备正式写入能力；AI 在最终交付前还必须重跑 `doctor`，确认授权码、路由、COM 和定时任务均达到相应就绪状态。

## 安装后续自动更新

第一次批准写入成功后，AI 使用首次配置时已经由用户确认的 `schedule` 安装 Windows 自动更新任务：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule install
```

任务使用计划任务程序的交互模式，因此要求用户保持登录。每次运行自动读取邮件、验证、备份并写入正式工作簿；内部临时验收副本在运行结束后删除，不生成需要用户打开的预览。无新数据时不写文件，冲突或异常时停止并记录错误，原表保持不变。

任务名包含运行实例 ID。重复安装只会替换本运行实例记录的任务。删除任务：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule remove
```

使用 `navctl.py schedule status` 查看任务、上次/下次运行、最近一次写入结果和备份路径。普通用户也可双击根目录的 `查看状态.bat`；失败详情保存在根目录 `logs/update-YYYYMMDD.log`，AI 排查即可。

不要为 UNC/网络路径中的运行目录安装任务，也不要复制其他电脑的 Python 路径或任务定义。

删除运行目录前，先移除任务和 DPAPI 密钥：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule remove
.\.venv\Scripts\python.exe navctl.py secret remove
```

备份和日志保留数量由 `app/config.json` 限制。

## 升级已部署环境

不得直接覆盖旧运行目录。先移除旧计划任务，再用最新版 `scripts/bootstrap.py` 在新的短路径创建运行目录；以新生成的 `app/config.json` 为基础迁移 AI 已复核的配置字段，把本地解析器逐个复核后迁入 `app/parsers/`，并保留新的 `runtime_id`。重新执行 `secret launch` 并用 `secret status` 复查，再完成离线演练、验证和一次预览，成功后安装新任务并停用旧目录。不要复制 `app/.venv`、密钥、`plan.json`、`automation-approval.json`、`run.lock`、预览、日志、备份或状态文件。

## 交给其他本地 AI

本仓库的 `SKILL.md` 遵循 Codex 与 Claude Code 使用的 Agent Skills 结构。对于 Cursor 或其他本地 Agent，让它完整阅读本文件及 [references/portable-use.md](references/portable-use.md)，并调用仓库内确定性的脚本，不要让它重新发明整套流程。

发布任何改动前运行：

```powershell
python -X utf8 scripts/privacy_audit.py
python -X utf8 scripts/selftest.py
python -X utf8 scripts/selftest.py --com
python -X utf8 scripts/package_check.py
```

`--com` 仅在装有 Excel/WPS 的 Windows 上运行，用临时虚构工作簿验证正式写入和公式缓存值。所有自测都会分阶段报告内容并自动清理临时文件。如果目标 AI 环境提供官方 Skill 校验器，也要一并运行。
