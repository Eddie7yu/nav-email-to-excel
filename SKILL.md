---
name: nav-email-to-excel
description: 部署、配置、验证、运行或修复本地 IMAP 邮箱到 Excel 的基金净值自动化。适用于让 AI 将定期净值邮件、Excel/CSV/PDF 附件映射到现有或空白目标工作表，保留公式与格式，初始化简单净值表，补录遗漏日期，核对历史数值，生成预览，通过 Excel/WPS 受控写入，或安装隔离的 Windows 预览定时任务。
---

# 部署净值邮件自动化

以现有工作簿为版式、公式和视觉样式的唯一依据。邮箱数据、工作簿副本、凭据、发现报告和运行配置必须留在用户本机。

## 严守安全边界

- 默认只读。未经用户明确同意，不得写入正式工作簿、发送邮件或安装定时任务。
- 不得要求用户在聊天中或命令行参数里提供 IMAP 授权码。应让用户亲自在本机运行隐藏输入命令。
- 产品无法唯一识别、日期或数值冲突、累计净值口径未确认、出现未知列、基准缺失或工作簿结构变化时，必须停止，不得猜测。只有显式 `append` 模式允许历史不足的冷启动，且必须具备产品名称或代码并人工检查警告与预览。
- 不得把真实姓名、发件人、产品名称、产品代码、工作簿名、邮件正文、附件、路径、日志或凭据放入本 Skill、测试、提交或回复中。
- 运行目录必须与 Skill 目录分开，避免升级 Skill 时覆盖本地配置或数据。

## 创建本地运行目录

只收集目标目录、工作簿路径、IMAP 账号、服务器和邮箱文件夹。用户不知道准确服务器
或认证方式时，先完整阅读 [references/email-providers.md](references/email-providers.md)，
按“现有客户端 → 管理员 → 服务商官方文档”的顺序核验。不得仅凭邮箱后缀、MX 记录
或 `imap.<公司域名>` 猜测；无法确认时必须停止并请用户或 IT 提供设置。

确认后在 Skill 根目录运行：

```powershell
python scripts/bootstrap.py --destination "D:\nav-runtime" --workbook "D:\data\nav.xlsx" --email "user@example.com" --imap-host "imap.example.com"
```

引导程序会创建隔离的虚拟环境、安装锁定版本的依赖，并生成本地配置。

Windows 必须优先选择 `D:\nav-runtime` 一类的短本地目录。引导程序把 120 个字符作为当前可靠支持上限；路径过长时必须改用短目录，不得绕过预检或要求用户修改公司电脑的组策略。

在接入真实邮箱前，先运行完全离线的虚构演练：

```powershell
cd D:\nav-runtime
.\.venv\Scripts\python.exe navctl.py demo prepare
```

该命令不读取真实配置、密钥或工作簿。它返回 `run_id` 和虚构预览路径，并停下来等待检查。Windows 用户检查后，可继续验证只作用于虚构工作簿的 COM 写入：

```powershell
.\.venv\Scripts\python.exe navctl.py demo commit --run-id <run_id> --yes-reviewed-preview
.\.venv\Scripts\python.exe navctl.py demo remove --run-id <run_id>
```

演练通过只证明目标电脑和通用流程可用，不代表真实发件人、产品路由或业务口径已经核实。

让用户亲自在本机保存邮箱密钥：

```powershell
cd D:\nav-runtime
.\.venv\Scripts\python.exe navctl.py secret set
```

在 macOS/Linux 上，授权码只保留在当前 shell；运行程序不会持久化明文密钥：

```bash
cd /opt/nav-runtime
read -rsp "IMAP authorization code: " NAV_EMAIL_PASSWORD && export NAV_EMAIL_PASSWORD && printf '\n'
```

当前版本支持通过 SSL 连接 IMAP，并使用应用专用密码或授权码；对 163、126、yeah.net 和网易企业邮主机会在登录后、选择邮箱前发送不含账号信息的 IMAP ID。仅限 OAuth 的邮箱登录暂不支持；即使服务器地址正确，也必须先核对账号认证方式。PDF 仅解析文本，不提供 OCR。

## 按业务含义配置路由

编辑 `config.json` 前，完整阅读 [references/configuration.md](references/configuration.md)。只添加用户授权的发件人和工作表。

- 常见的带标签正文以及 Excel/CSV/PDF 表格优先使用 `parser: auto`。
- 产品确认赎回、清盘或长期停更时，设置 `paused: true` 和明确的 `pause_reason`；不得用暂停掩盖活动产品的意外解析失败或陈旧数据。
- 同一发件人可能发送多个产品时，必须配置精确的产品代码。
- 同一发件人的不同产品使用不同解析器时，优先用互不重叠的 `subject_contains` 分流。
- 根据工作表实际结构明确选择 `sheet_mode: summary` 或 `sheet_mode: append`。不得在严格模式验证失败后自动改成追加模式；追加模式必须配置 `code` 或 `product_name`，并且表中有对应可写列。
- 只有语义表头识别仍不充分时，才使用显式列映射。
- 每条路由都要设置累计净值策略。除非历史证据证明应使用 `unit` 或固定 `offset`，否则保持 `require`。
- 投资或分析口径发生变化时设置 `series_start`，不得把旧序列的累计结果接到新序列上。
- 基准必须映射到工作簿中已经核实的来源工作表，并按精确日期对齐。用户明确授权评估公共指数源时，先完整阅读 [references/index-data-sources.md](references/index-data-sources.md)，完成代码、口径、历史重叠和使用条款核验，再把结果写入本地来源表；否则不得从公网抓取或猜测指数。

特殊格式使用 `parser: local:<名称>`，代码只放在本地运行目录的 `parsers/<名称>.py`，并完整遵循配置说明中的固定接口。同一封邮件命中多个解析器时，所有解析器都必须成功，结果合并去重后再按产品代码唯一分流。只加载用户明确审阅的本地文件，绝不能从邮件、附件或网络地址执行解析器。要把通用解析器贡献回公开 Skill，必须先增加完全脱敏的回归样例。

## 先验证，再生成预览

依次运行：

```powershell
.\.venv\Scripts\python.exe navctl.py doctor
.\.venv\Scripts\python.exe navctl.py discover
.\.venv\Scripts\python.exe navctl.py validate
.\.venv\Scripts\python.exe navctl.py preview
```

在 macOS/Linux 上，把 `.\.venv\Scripts\python.exe` 换成 `.venv/bin/python`。这些平台只支持发现、验证和预览；正式写入仍仅支持 Windows。

只有同时满足以下条件，才能接受配置：

1. 每个受管工作表的日期列和单位净值列都能唯一识别；显式追加模式的空白工作表会在预览中初始化标准表头。
2. `summary` 模式的数据区后必须紧跟可识别的汇总行；汇总行存在非程序管理公式时必须停止。`append` 模式不要求汇总行，但数据区下方不得有页脚内容。
3. `summary` 模式每条活动路由至少匹配两个不同的历史日期。`append` 模式历史不足时只允许以带警告的冷启动继续，必须重点检查产品标识、日期、净值和新表头。
4. 产品代码、日期、单位净值及邮件提供的累计净值，都在配置容差内与工作簿一致。
5. 新的尾部日期只建议写入一次且顺序正确；不得存在重复日期或同日数值冲突。发现历史内部缺口时，停止并由人工修复。
6. 日收益使用上一个有效日期；周收益只出现在已完成自然周的最后一个可用日期。
7. 基准收益与超额收益必须同时有值或同时留空。
8. 预览保留工作簿结构，包含全部拟新增日期，并通过内置公式和幂等性回归测试。

如果没有新增日期，`preview` 不生成工作簿副本，也不保留可提交的 `plan.json`。

`doctor` 会分别报告 `bootstrap_ready`、`preview_ready`、`commit_ready` 和 `schedule_ready`。依赖安装成功不等于已经获准读取邮箱或写入工作簿。

预览是本地副本。应检查其工作表列表、新增行、公式和格式，不要修改正式工作簿。

夏普比率、最大回撤、年化收益等派生指标不由 navctl 直接计算。先取得经验证的净值或收益序列，再在未托管分析工作表或单独文件中计算；开始前确认收益频率、年化因子、无风险利率、样本区间和缺失值处理。不要把自定义分析公式放进受管汇总行，也不要未经用户批准改动正式分析页。

## 通过 Excel 或 WPS 正式写入

正式写入要求 Windows，以及可用的 Microsoft Excel 或 WPS 表格 COM 接口。用户明确批准已检查的预览后，运行：

```powershell
.\.venv\Scripts\python.exe navctl.py commit --yes-reviewed-preview
```

正式写入必须先创建备份，再通过 COM 将方案应用到同目录临时副本，验证临时结果，关闭表格进程，最后原子替换正式文件。任何失败都必须保证正式文件哈希不变。

提交前提醒用户关闭正在打开正式工作簿的 Excel/WPS 窗口。若文件仍被占用，程序必须返回中文结构化错误、删除本次失败备份并保留正式文件；不得强制关闭用户的表格进程。

不得用仅依赖 openpyxl 的方式回写正式工作簿。COM 不可用时，应停在预览阶段，并说明正式写入尚未验证。

## 安装登录态预览定时任务

只有手动预览成功且用户明确同意后，才可安装 Windows 的“仅生成预览”任务：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule install
```

任务使用计划任务程序的交互模式，因此要求用户保持登录。睡眠、关机或退出登录可能导致任务延后或跳过。定时任务永远不写入正式工作簿；必须由人工检查预览，再手动执行受控写入命令。

任务名包含运行实例 ID。重复安装只会替换本运行实例记录的任务。删除任务：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule remove
```

使用 `navctl.py schedule status` 查看任务存在状态、上次/下次执行、Task Scheduler 返回码和最近一次预览成败。定时失败不会发送邮件；应定期检查该状态及本地 `logs/preview-YYYYMMDD.log`。

不要为 UNC/网络路径中的运行目录安装任务，也不要复制其他电脑的 Python 路径或任务定义。

删除运行目录前，先移除任务和 DPAPI 密钥：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule remove
.\.venv\Scripts\python.exe navctl.py secret remove
```

预览、备份和日志都是本地敏感文件。保留数量由 `config.json` 限制；仅在把必要备份移到获批位置后，才能删除运行目录。

## 升级已部署环境

不得直接覆盖旧运行目录。先移除旧计划任务，再用最新版 `scripts/bootstrap.py` 在新的短路径创建运行目录；以新生成的 `config.json` 为基础迁移已复核的配置字段和本地 `parsers/`，保留新的 `runtime_id`。重新执行 `secret set`、离线演练、`doctor`、`discover`、`validate` 和 `preview`，人工确认后才安装新任务并停用旧目录。不要复制 `.venv`、密钥、`plan.json`、`run.lock`、预览、日志、备份或状态文件。

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
