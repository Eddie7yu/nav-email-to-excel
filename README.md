# 净值邮件写入 Excel Skill

这是一个面向中文用户和本地 AI Agent 的公开 Skill：它把“从 IMAP 邮箱读取净值邮件、核对历史数据、生成 Excel 预览、经人工批准后受控写入”的流程固化成可验证的脚本。

它不是一键黑盒，也不会替用户猜业务口径。首次部署仍需用户确认发件人、产品代码、目标工作表、累计净值规则、序列起点和基准来源；邮箱授权码只由用户在本机隐藏输入。

## 它解决什么问题

- 从邮件正文或 Excel、CSV、文本型 PDF 附件中提取净值。
- 把发件人和产品精确映射到已有工作表，不靠模糊猜测。
- 在处理新数据前，先用至少两个历史日期验证产品代码、日期和净值。
- 只追加工作簿尾部的新日期；发现历史内部缺口、重复日期或同日冲突时停止。
- 保留原工作簿结构、公式和样式，生成可检查的本地预览。
- 正式写入前绑定配置、正式工作簿和预览的哈希，防止“检查的文件”和“写入的文件”不是同一版本。
- Windows 下通过 Excel/WPS COM 写入临时副本，验证后备份并原子替换正式文件。
- 定时任务只生成预览，永远不会无人值守地写入正式工作簿。

## 为什么能少踩很多坑

仓库把此前最容易反复出错的环节做成了固定规则和回归测试，包括：

- 依赖版本锁定、隔离虚拟环境、Windows 短路径预检和可读的 Unicode 错误；
- 邮件发件人二次核对、产品代码精确匹配、停更路由显式暂停和陈旧数据拦截；
- 累计净值口径、序列重置、日/周收益和基准对齐；
- 未来日期、异常净值跳变、重复写入和历史缺口保护；
- 汇总公式失败关闭、格式、工作表结构、预览完整性及正式文件不变性检查；
- 本地密钥、日志、预览、备份的隔离与公开仓库脱敏审计。

它不能保证从未见过的邮件格式都自动兼容，也不能替代对业务口径的人工确认。遇到无法证明的情况，正确行为是停止并给出原因；如确需新增特殊解析器，可以使用运行目录中的受信任 `parsers/` 扩展点，但必须先用完全脱敏的样例回归，绝不能从邮件或附件加载代码。

第一次部署可先阅读面向非程序员的[中文使用流程](使用流程.md)。

## 适用范围与限制

- 邮箱：IMAP over SSL，使用应用专用密码或授权码；支持标准 IMAP，以及需要登录后发送 ID 的 163、126、yeah.net 和网易企业邮主机；暂不支持仅限 OAuth 的登录。常见服务商地址、认证限制和核验方法见 [IMAP 服务商参考](references/email-providers.md)。
- 附件：Excel、CSV 和文本型 PDF；扫描 PDF 暂不提供 OCR。
- 预览：Windows、Linux、macOS 均可运行 Python 流程；持续集成覆盖 Windows、Linux，并对 macOS 做预览 smoke test。
- 正式写入：仅支持 Windows，并要求 Microsoft Excel 或 WPS 表格的 COM 接口可用。
- 定时任务：仅支持 Windows 登录态，且只生成预览。
- 工作簿：面向已有 `.xlsx`/`.xlsm`；受管数据区后必须紧跟可识别的汇总行。纯追加、没有汇总行的表暂不支持；无法证明可安全扩展的汇总公式会停止预览。

## 给本地 AI 的最快用法

以 Codex 为例：

```powershell
git clone https://github.com/Eddie7yu/nav-email-to-excel "$env:USERPROFILE\.codex\skills\nav-email-to-excel"
codex
```

然后输入：

```text
请使用 $nav-email-to-excel，为这个工作簿部署一套预览优先的本地运行环境：<工作簿绝对路径>
```

Claude Code、Cursor 及其他本地 AI 的安装方法和完整提示词见 [交给其他本地 AI 使用](references/portable-use.md)。AI 应调用仓库内脚本，不应自行重写核心逻辑。

如果在 Windows PowerShell 中手动查看中文说明，应明确指定 UTF-8：

```powershell
Get-Content -Raw -Encoding utf8 SKILL.md
```

## 先做一次完全离线演练

离线演练用于验收目标电脑，不会读取真实 `config.json`、邮箱、密钥或正式工作簿。它使用虚构邮件和虚构 Excel，实际走过邮件发现、两个历史日期核验、预览、公式和保护逻辑，从而把“电脑环境问题”和“真实业务配置问题”分开。

在已经创建的运行目录中先生成虚构预览：

```powershell
.\.venv\Scripts\python.exe navctl.py demo prepare
```

命令会分阶段报告结果，返回 `run_id` 和虚构预览路径，然后停下来等待检查。Windows 用户检查预览后，可以仅对虚构工作簿演练 Excel/WPS COM 写入：

```powershell
.\.venv\Scripts\python.exe navctl.py demo commit --run-id <run_id> --yes-reviewed-preview
```

完成后清理该次虚构演练：

```powershell
.\.venv\Scripts\python.exe navctl.py demo remove --run-id <run_id>
```

`demo commit` 不会触碰正式工作簿。macOS/Linux 可以完成 `demo prepare`，但不能演练仅限 Windows 的 COM 正式写入。

## 首次部署大致流程

1. AI 完整阅读 [SKILL.md](SKILL.md) 和[配置说明](references/configuration.md)。
2. 用户提供本机目标目录、工作簿路径、邮箱账号、IMAP 服务器和邮箱文件夹；不知道服务器时，AI 按 [IMAP 服务商参考](references/email-providers.md) 核验，不能凭邮箱后缀猜测。
3. 运行 `scripts/bootstrap.py`，创建独立运行目录和锁定依赖的虚拟环境。
4. 运行 `navctl.py demo prepare`；Windows 可继续完成虚构 COM 写入演练。
5. 用户亲自在本机运行 `navctl.py secret set`，隐藏输入邮箱授权码。
6. AI 在本机发现工作簿结构，并根据用户确认的信息配置路由和业务口径。
7. 依次执行 `doctor`、`discover`、`validate`、`preview`。
8. 用户检查预览中的工作表、新增行、公式和格式。
9. 只有用户明确批准后，Windows 才执行 `commit --yes-reviewed-preview`；否则正式工作簿保持不变。

Windows 运行目录请优先使用 `D:\nav-runtime` 之类的短本地路径。引导程序把 120 个字符定义为当前支持上限，并会在创建文件前拒绝更长路径；这是本工具的可靠部署边界，不是 Windows 的通用路径上限。

已安装定时预览后，使用以下命令查看最近一次运行结果，并定期检查本地 `logs/`：

```powershell
.\.venv\Scripts\python.exe navctl.py schedule status
```

## 升级已部署的运行环境

不要在旧运行目录中覆盖脚本、替换 `.venv` 或直接复制新版模板。安全升级采用“新建运行目录、迁移配置、重新验收”，旧目录在验收完成前保留为回退副本：

1. 在旧运行目录执行 `navctl.py schedule remove`，避免新旧任务同时运行。
2. 更新干净的 Skill 仓库，在新的短路径重新运行 `scripts/bootstrap.py`；正式工作簿路径保持不变。
3. 以新生成的 `config.json` 为基础，人工迁移并复核 `routes`、`column_overrides`、`style`、`schedule`、`validation` 和 `retention`。不要用旧文件整体覆盖，以免带回旧的结构版本或 `runtime_id`。
4. 如有本地 `parsers/`，逐个审阅后复制到新运行目录；不要迁移 `.venv`、密钥、`plan.json`、`run.lock`、预览、日志或定时状态文件。
5. 在新目录重新执行 `secret set`、离线演练、`doctor`、`discover`、`validate` 和 `preview`。只有新预览经人工确认后，才重新安装定时任务并停用旧目录。

当前没有 `navctl upgrade`，这是有意保留的失败关闭设计：配置结构或运行时依赖变化时，升级必须经过重新验收，不能静默改动正式环境。

## 公开仓库与本地数据的边界

公开仓库只包含通用 Skill、确定性脚本和虚构测试数据。真实邮箱、发件人、产品、工作簿、路径、日志、预览、备份、配置和凭据必须留在单独的本地运行目录，绝不能提交或分享。

发布前应运行：

```powershell
python -X utf8 scripts/privacy_audit.py
python -X utf8 scripts/selftest.py
python -X utf8 scripts/selftest.py --com
python -X utf8 scripts/package_check.py
```

普通 `selftest.py` 适用于所有平台；`--com` 会在 Windows 上额外创建临时虚构工作簿，真实调用 Excel/WPS COM，复核公式缓存值后自动清理。测试输出会逐阶段说明检查内容，全程不使用真实资料。

## 许可证

[MIT License](LICENSE)
