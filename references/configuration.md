# 运行配置说明

## 目录

- [如何使用本说明](#如何使用本说明)
- [顶层字段](#顶层字段)
- [确定 IMAP 服务器](#确定-imap-服务器)
- [自动发现发件人和产品](#自动发现发件人和产品)
- [普通用户新增产品的两种入口](#普通用户新增产品的两种入口)
- [路由字段](#路由字段)
- [工作表模式](#工作表模式)
- [本地解析器](#本地解析器)
- [列识别](#列识别)
- [基准映射](#基准映射)
- [派生专业指标](#派生专业指标)
- [定时任务](#定时任务)
- [零新增首次验收与错误反例](#零新增首次验收与错误反例)
- [本地敏感文件](#本地敏感文件)

## 如何使用本说明

本说明记录配置接口、经过验证的默认方案和已知错误案例，不是外部邮箱与工作簿问题的穷举，也不要求 Agent 严格照章节顺序排错。场景与前提一致时优先复用现有方案；前提不同或没有现成案例时，可以在本机自主检查邮件和工作簿证据、操作副本、编写 `app/diagnostics/` 临时工具、比较其他受信任解析方法，并根据结果选择通用修复、本地解析器或明确配置。

探索方法可以不同，但配置语义和生产验收不能靠猜测：真实数据不得进入公开仓库；产品、日期、净值、频率和累计口径必须有证据；正式结果仍通过 `navctl.py preview`、首次批准、验证、备份和受保护写入。临时原型不得直接成为正式写表入口或计划任务，交付前应清理。

## 顶层字段

`app/config.json` 由本地运行实例生成，绝不能提交到 Git。普通用户不需要打开或修改它；AI 从运行目录的 `app/` 执行以下命令。

| 字段 | 含义 |
| --- | --- |
| `schema_version` | 配置结构版本，当前为 `1` |
| `runtime_id` | 随机的部署实例标识，用于隔离密钥和定时任务 |
| `workbook_path` | 正式 `.xlsx` 或 `.xlsm` 工作簿的绝对路径 |
| `workbook_mode` | `existing` 表示使用已有表；`bundled-template` 表示由 `--new-workbook` 登记、等待或已经完成模板初始化 |
| `imap` | 服务器、端口、用户、邮箱文件夹、回看天数、过滤前邮件头上限、目标邮件大小和数量限制；不得包含密码 |
| `routes` | 已授权的“发件人/产品 → 工作表”映射 |
| `sheet_reviews` | 可选的未接管 Sheet 覆盖审计，由 AI 根据本地证据维护，普通用户不填写 |
| `column_overrides` | 可选的逐工作表语义列映射 |
| `style` | 已有表使用 `infer` 保留原样；内置模板使用 `cn-red-up-green-down`，正收益红、负收益绿、接近零黑 |
| `schedule` | Windows 正式部署必填的自动更新时间；仅在引导尚未完成时允许为空 |
| `validation` | 历史样本数量和数值容差 |
| `retention` | 本地备份、预览最大数量及日志保留天数 |

网易 163、126、yeah.net 和网易企业邮的 IMAP 主机要求在登录后、选择邮箱文件夹前发送 ID。运行时会按主机名自动完成该握手，ID 只包含程序名和主版本，不包含邮箱账号、路径或系统信息；其他 IMAP 主机不发送该扩展命令。

### 工作簿逐页覆盖审计

`products status` 和 `products sync` 会输出 `workbook_coverage`，逐页归入活动路由、基准待审路由、暂停路由、参考页、明确排除、邮箱无证据、需本地解析器、业务待确认或未分类。AI 对未接管页完成只读核实后，可在本地配置记录：

```json
{
  "sheet_reviews": {
    "示例归档产品": {
      "status": "excluded",
      "reason": "本地资料已证明不再持续更新"
    },
    "示例特殊公告产品": {
      "status": "local_parser_required",
      "reason": "正式通知需要经过脱敏验证的供应商专用解析器"
    },
    "示例待确认产品": {
      "status": "business_review",
      "reason": "持续正式来源尚不能由本地证据唯一确认"
    }
  }
}
```

可用状态为 `excluded`、`no_mail_evidence`、`local_parser_required`、`business_review`。理由必须是单行、非公式、最多 200 字，只存在本机；不得把真实产品身份或邮件内容写进公开仓库。一个 Sheet 建立路由后，产品命令会自动移除其旧审计项。`all_sheets_classified` 表示盘点数量闭环，不表示所有页都已自动化；`action_required_sheets` 会继续计入基准待审、解析器待补、业务待确认和未分类项。

## 确定 IMAP 服务器

用户不知道服务器或认证方式时，完整阅读 [IMAP 服务商与核验方法](email-providers.md)。
公司邮箱后缀不能证明其托管服务商，MX 记录也不是 IMAP 地址；不得把 MX 主机直接写入
`imap.host`，不得未经确认猜测 `imap.<公司域名>`。优先采用现有邮件客户端、单位管理
员或服务商当前官方文档提供的设置。仅支持 OAuth 的账号即使主机和端口正确也不能使
用当前运行时。

## 自动发现发件人和产品

用户保存授权码后，优先由 AI 运行只读头部预筛选：

```powershell
.\.venv\Scripts\python.exe navctl.py propose --headers-only --header-limit 25
```

它只读取最近指定数量邮件的 `From/Subject`，不查询大小、不下载正文或附件，并把本地候选写入 `app/route-proposal-headers.json`。AI 将其与工作簿中的产品名称、代码和 Sheet 语义比较；样本不足时可以按明确诊断目的扩大 `--header-limit`，不得无边界扫描。下一次完整解析至少要从该报告选择一个精确发件人。

能确定精确发件人和稳定主题片段时，运行：

```powershell
.\.venv\Scripts\python.exe navctl.py propose --sender "sender@example.com" --subject-contains "稳定主题片段"
```

若 Windows 外部 Agent 所在终端不能可靠传递中文参数，在 `app/diagnostics/` 建立 UTF-8 参数文件，每行只写一个完整参数：

```text
propose
--sender
sender@example.com
--subject-contains
示例产品净值通知
```

从 `app/` 运行：

```powershell
.\.venv\Scripts\python.exe navctl.py @diagnostics\propose-args.txt
```

UTF-8 BOM 可以保留；空行和以 `#` 开头的说明会忽略。参数值不加 PowerShell/cmd 引号，嵌套 `@` 文件、非 UTF-8、NUL、替换字符或明显的错误解码会在连接邮箱前停止。成功报告的 `scan.filter_input_encoding` 会记录输入来源和编码检查结果，因此 `headers_matched: 0` 只有在编码检查通过后才可解释为该筛选范围真实零命中。参数文件只存本机并在诊断后清理。

程序先让 IMAP 服务端按精确发件人缩小整个回看期结果，只读取该范围的最小邮件头；主题片段在本地复核，同时命中的邮件才查询大小、下载完整 MIME 并解析。报告中的 `server_since_matches`、`server_sender_matches`、`headers_fetched` 和 `messages_fetched` 可用来判断范围是否真正收窄。程序仍会解析并精确核对 `From`，不能把服务端搜索当成身份认证。产品身份只存在于正文或附件、主题无法可靠收窄时，省略主题但保留发件人：

```powershell
.\.venv\Scripts\python.exe navctl.py propose --sender "sender@example.com"
```

CLI 的完整候选解析默认使用最多 30 天、每批 50 封，并为下载与解析阶段分别设置 120 秒软时间预算；主题存在时服务端搜索也包含主题条件，本地仍二次核验。运行中 stderr 与 `app/route-proposal-progress.json` 会显示当前阶段、邮件进度和附件累计数。若输出 `partial: true`，本批结果只保存在 `app/route-proposals.partial.json`，不能用于接管；运行 `navctl.py propose --resume` 从 `resume_before_uid` 继续。需要更早历史时用 `--lookback-days` 明确扩大同一产品范围，仍按批次恢复。完成后 `app/route-proposals.json` 才会被原子更新，部分检查点自动删除。

搜索、最小邮件头、大小或单封正文读取遇到瞬时 IMAP 会话断开时，客户端最多自动重连两次并重试当前无副作用步骤。每次重连都重新登录、完成服务商握手、只读选择同一文件夹并比较 `UIDVALIDITY`；不一致立即停止，防止把旧 UID 用到已重建的邮箱。候选报告和 `route-report.json` 的 `imap_reconnects` 记录次数，候选进度还会出现 `reconnect` 阶段。超过重连次数后不发布本轮不完整结果；选择式候选使用既有 UID 检查点恢复。原始 MIME/附件不作为跨运行缓存长期保存，避免敏感资料落盘和陈旧附件被误复用；只有完整 discovery 成功后，才保存受邮箱、路由、解析器和源码指纹约束的净值行快照。

新运行目录的日常 `imap.lookback_days` 默认 60 天，避免每次预览和计划任务重复下载半年邮件；它应覆盖产品实际披露周期、节假日缓冲和足够历史重叠。首次历史补录不受这个日常窗口上限约束：先对单一产品用 `propose --lookback-days 180` 等受控批次完成身份核验，再用 `preview --lookback-days 180` 生成一次性历史预览。该参数不会改变配置或后续计划任务窗口，但会进入本次方案；用户批准后正常 `commit`。不得为了补一次历史把长期计划任务永久留在过大的窗口。

普通无参数 `propose` 会在连接邮箱前停止。只有扩大头部样本仍无法确定候选发件人、Agent 已核实 `lookback_days`、`max_messages`、`max_message_bytes` 和 `max_total_bytes`，并且确有受控诊断需要时，才显式运行 `propose --allow-unscoped-full-scan`。完整候选报告包含可解析候选的真实发件人、主题样例、产品代码、净值日期/数值和近期到达时间，只存在于本地运行目录。AI 应读取该报告；已有工作簿时再读取其结构，自动生成 `routes`、`column_overrides` 和 `schedule`。没有工作簿时，AI 从邮件日期规律判断日频/周频，只有证据冲突才询问。不得要求用户手工抄写发件人邮箱、产品代码或逐项填写列号。

主题代码唯一、附件数据行没有代码时，可以显式运行：

```powershell
.\.venv\Scripts\python.exe navctl.py propose --sender "sender@example.com" `
  --subject-contains "DEMO01" --subject-product-code "DEMO01"
```

程序只在两个代码标准化后完全相等、每封邮件得到单一无冲突序列且没有混合行代码时绑定。`subject_code_binding`、候选 observation 来源和后续 `route-report.json` 都会记录结构化证据；产品接管会继承 `subject_product_code`。这不是 `allow_sender_only`，也不能用于宽泛主题。

若 `headers_matched > 0` 但 `parsed_records == 0`，检查 `parse_gap_summary`。其中只包含正文类型是否存在、附件 MIME 类型计数、受控失败类别和是否建议本地解析器，不含完整主题、正文或附件名。先用脱敏样例验证通用/本地 parser，再重新生成完整候选。

头部预筛选和紧接的一次选择式完整解析属于同一轮候选发现。成功生成完整候选报告后不要再次重复运行 `propose`。完成路由配置后直接运行一次 `preview`；它会在同一次授权邮件扫描结果上完成发现、历史验证和预览。`discover`、`validate` 仅用于分阶段诊断，不应在每次 `preview` 前固定串行执行。

完整 discovery 成功时，运行时把最终路由行和脱敏报告写入 `app/discovery-snapshot.json`。该文件只在本机使用，不进入公开仓库，也不替代邮箱校验。后续如果失败发生在工作簿列映射、基准来源、数组公式或预览构建阶段，可运行：

```powershell
.\.venv\Scripts\python.exe navctl.py preview --reuse-discovery
```

命令仍连接邮箱，但只读取当前授权范围的轻量邮件头、UID 和 `RFC822.SIZE`，不下载正文或附件。程序同时比较 IMAP 范围、活动路由的发件人/主题/排除词/代码/别名/解析器/`series_start`、核心解析源码和本地解析器哈希；全部一致才恢复快照中的净值行，然后重新读取当前工作簿、验证历史并生成新的完整预览。邮箱新增或改动、路由或解析器变化、临时回看窗口、局部 `--sheet` 都会拒绝复用，Agent 改用普通 `preview`。快照复用只优化后段迭代，不允许跳过跨路由检查、当前工作簿哈希、计划签名或首次用户验收。

候选与工作表能通过产品代码或产品名称任一唯一匹配时直接配置并生成首次预览；历史重叠用于增强核验和消除警告，不得因为不足两个日期而阻止预览。只有多个映射同样合理、累计净值规则无法从本地证据证明，或一个发件人包含多个无代码产品且主题也无法分流时才询问用户。

## 普通用户新增产品的两种入口

用户只需表达“接管我建好的 Sheet”或“格式照现有 Sheet”，无需理解内部模式。以下命令是这两种入口的已验证默认实现；特殊工作簿可以采用其他本地诊断和兼容方式，但不得增加用户配置负担，也不能绕开首次预览和正式写入保护。部署后优先运行 `navctl.py products sync --use-existing-proposals` 复用刚生成的完整候选报告；报告不再适用时才用 `products sync --sender ...` 做选择式刷新。无范围刷新会被拒绝，除非 Agent 对特殊情况显式使用 `--allow-unscoped-full-scan`。报告包含：

- `new_candidates`：尚未配置的发件人/产品代码候选；
- `alias_pending`：邮件代码与某个同发件人主代码只有结构化份额后缀差异，必须核实后保存精确别名，不能模糊接管；
- `matched_candidates`：已配置且仍在回看范围内出现的产品；
- `configured_not_seen_in_lookback`：本次回看没有发现的已配置产品，只是诊断线索，不会自动暂停；
- `workbook_missing_sheets`：配置引用但工作簿缺失的页面；
- `unmanaged_workbook_sheets`：未被路由或基准使用的页面，可能是分析页，不应自动删除。
- `review_required`：仍处于基准/超额审查状态的活动产品数量；每条路由的 `review_issue` 区分 `benchmark-source-unresolved` 与 `benchmark-license-unresolved`。大于零时可以预览，但不能提交或上线自动更新。

离线复查上一次候选报告可使用 `products sync --use-existing-proposals`；只看当前状态使用 `products status`。用户不需要看报告或候选序号，由 AI 完成匹配。

### A. 用户已经建立目标 Sheet

命令：

```powershell
.\.venv\Scripts\python.exe navctl.py products adopt `
  --proposal-index 1 --sheet "用户新建的 Sheet"
```

识别顺序：

1. 以工作表的语义表头定位产品代码、产品名称、单位净值、累计净值、日期和收益列；
2. 有紧邻累计/合计行时判为 `summary`，无页脚时判为 `append`；用户不选择内部模式；
3. 优先用 Sheet 中代码与邮箱候选的交集唯一定位产品；其次使用候选中唯一代码；
4. 日收益表头判为日频，周收益表头判为周频；表头不明确时读取已有真实日期间隔；
5. 邮件提供累计净值时使用 `require`；邮件不提供时，只有工作簿历史能证明累计净值始终等于单位净值才使用 `unit`，否则停止并报告真实业务歧义。

`adopt` 只修改本地 `app/config.json`，正式工作簿字节必须完全不变。选择式 `propose` 的精确发件人和主题范围会随候选自动继承；Agent 不要重复传 `--subject-contains`，更不能在接管时悄悄扩大范围。若确需改变范围，先按新范围重新运行 `propose`。仅在候选含多个代码但 Sheet 无法唯一映射时传 `--code`；只有 Sheet 未写完整名称时才传 `--product-name`。旧的无范围候选报告确实需要主题分流时，才显式传 `--subject-contains`。

多产品首次部署先运行同一命令并增加 `--inspect-only`。它只返回推断布局、频率、累计策略、`column_roles`、建议 `series_start` 和 `review_required`，配置与工作簿哈希都不变。`column_roles` 同时显示 Excel 列字母和原始表头，Agent 应在预览前核对具体指数名称是否被识别为基准点位/收益，而不是等公式门禁才发现错列。

历史范围是 Agent 面向不同工作簿选择的内部策略，不是普通用户配置项：

- `--history-scope tail` 是默认续接方式。已有真实历史的 Sheet 建议“原表最后日期 + 1 天”，并用最后既有点只读核验；空表或严格预留行不套用该边界。
- `--history-scope mail-history` 不自动设置表尾边界，让可唯一证明的邮箱历史进入首次受保护预览。用户要求补历史、工作簿明确缺点或本地证据表明旧区间需要复核时使用；历史内部插入、值冲突和公式风险仍会停止，不能静默回填。
- 更特殊的投资起点或口径切换可由 Agent 在配置中设置明确 `series_start`。默认值不适用不等于关闭校验。

先正式接管所有 `ready_for_direct_adoption: true` 的产品，生成并批准第一份可提交预览；基准来源待核实的产品留到单独审查批次。安全阻断仍保留，但不让少数待审产品拖住证据完整产品。

确认份额级别代码确属同一产品后运行：

```powershell
.\.venv\Scripts\python.exe navctl.py products alias --sheet "用户新建的 Sheet" --code "邮件中的完整份额代码"
```

程序只接受当前同发件人候选报告中真实出现的完整代码，并写入 `code_aliases`。确认后必须重新预览；运行时只按主代码或别名完整相等路由，不使用模糊字符串。

目标页含空白的基准/超额预留列时正常接管，不把“有这个表头”等同于“已有基准业务”。如果这些列已有点位、数值或公式，而 AI 暂时无法证明可靠来源，程序自动增加内部 `benchmark_review_only: true`：允许验证净值并生成审查预览，但 `plan.json` 会包含 `committable: false` 和脱敏的 `blocking_reviews`，`commit` 必须拒绝，`doctor` 的 `write-rules-resolved`、`commit_ready`、`schedule_ready` 也保持失败。审查预览只在新增行和汇总行清空未确认的基准点位/收益/超额，已有历史与正式工作簿不变。AI 找到并核验技术来源后，以标准 `benchmark` 映射替换顶层状态；使用许可尚未确认时在该映射内保留 `review_only: true`，许可批准后再取消并重新运行 `preview`。不要让普通用户编辑这些字段，也不要把任何待审状态升级成“连预览都不能生成”。

### B. 照受管 Sheet 的格式新建

命令：

```powershell
.\.venv\Scripts\python.exe navctl.py products clone `
  --proposal-index 1 `
  --sheet "新产品" `
  --copy-from "参考产品" `
  --product-name "邮件中的完整产品名称"
```

执行边界：

- 参考页必须是同一工作簿中已受管的 `summary` 产品页；纯追加表先由用户或 AI 在 Excel/WPS 中复制好，再走 A 路径；
- 目标页必须不存在；关闭正在打开正式工作簿的 Excel/WPS 窗口；
- 程序在 `backups/` 留存写前备份，只对同目录临时副本执行 COM 复制和清空；
- 复制整张 Sheet 后删除参考产品历史和分段，清空产品说明及业务值/公式，保留一条“新代码/名称 + 邮箱最早真实日期 + 空净值”的预留行和一条干净累计行；
- 核验参考页及其他 Sheet 的值、公式和拓扑未变后才原子替换正式工作簿；配置保存失败时从备份恢复；
- 默认不继承基准。只有可靠资料证明新产品与参考页使用同一基准时才传 `--inherit-benchmark`。

这两种新增都会继承本轮选择式候选的主题范围，并使自动写入批准失效。若命令参数与已选范围冲突，程序拒绝接管并要求重新运行 `propose`。紧接着运行一次 `preview`，检查新产品的历史、日期频率、收益公式和样式，再批准 `commit`。原计划任务不用重装。

## 路由字段

```json
{
  "sender": "sender@example.invalid",
  "subject_contains": "NAV",
  "subject_product_code": null,
  "sheet": "Demo Fund",
  "sheet_mode": "summary",
  "code": "DEMO01",
  "code_aliases": ["DEMO01A"],
  "product_name": "Example Fund",
  "parser": "auto",
  "paused": false,
  "allow_sender_only": false,
  "cumulative_policy": "require",
  "cumulative_offset": null,
  "return_basis": "cumulative",
  "return_frequency": "weekly",
  "data_frequency": "auto",
  "series_start": "2026-01-01",
  "baseline_overlap": "last_existing_point",
  "max_staleness_days": 14,
  "benchmark": null
}
```

- 产品代码按不区分大小写的方式标准化，但标准化后必须完全相等。
- `code_aliases` 只保存 Agent 已从本地证据确认的完整邮件代码；同一发件人的主代码和所有别名必须全局唯一。不得自动截掉份额后缀或使用前缀/包含匹配。
- 产品代码必须保存为带引号的 JSON 字符串。系统拒绝数字类型，避免前导零消失。
- `sender`、`code` 和适用的 `subject_contains` 优先来自 `navctl.py propose` 的本地候选报告，不由用户手工录入。
- `subject_product_code` 仅用于“精确主题代码 → 同封邮件无代码净值行”的显式绑定，必须等于主代码并同时配置精确 `subject_contains`；路由报告和预览会显示绑定证据与邮件数。
- `subject_contains` 优先保存托管人正式净值通知的稳定主题前缀。核心运行时会在下载正文前排除虚拟/模拟估算，以及不含明确净值通知语义的交易确认、月报和季报，并在 `excluded_non_nav_reasons` 按 `virtual-estimate`、`transaction-confirmation`、`periodic-report` 等受控类别计数。主题同时含“净值”或独立 `NAV` 等明确信号时继续进入解析，避免把供应商的复合通知静默排除。
- `subject_excludes` 是 Agent 根据本地原始主题核实后保存的供应商特有排除片段列表，不交给普通用户填写。每项必须非空、单行且不与 `subject_contains` 相同；运行时先验证正式主题前缀，再应用排除片段，最后仍按代码/别名路由。不得加入“净值”“通知”一类会吞掉正式邮件的宽泛词。
- 若获准发件人也会发送非净值邮件，应配置 `subject_contains`。范围内的每一封邮件都必须成功解析；解析失败会阻止预览，不能静默跳过。解析成功后，带明确产品代码但不属于任何活动路由的行不属于受管数据，可以选择式忽略。
- 已配置路由的邮箱扫描先批量读取 `From`/`Subject` 邮件头，按同一发件人的全部活动路由主题过滤后，才应用数量和字节限制并下载完整邮件。`imap.max_header_messages` 默认 20000，限制过滤前单个授权发件人在回看期内的邮件头数量，且不得小于 `max_messages`；超限时停止并缩短回看窗口或经核实后调高。若该发件人的任一路由没有 `subject_contains`，则其全部邮件仍属于授权范围。
- 产品确认赎回、清盘或长期停更时，设置 `paused: true` 并填写不超过 200 字的 `pause_reason`。暂停路由不会连接对应发件人、不会参与陈旧检查或预览，并会出现在发现报告的警告中。所有路由都暂停时，发现必须失败；不得用暂停掩盖活动产品的意外异常。
- 邮件中没有稳定产品代码，但产品名称、专用发件人或互不重叠的主题范围已使路由唯一时，可设置 `allow_sender_only`。同一范围仍可能对应多个产品时不得使用。
- 累计净值必须来自邮件时，使用 `cumulative_policy: require`。
- 只有历史样本证明单位净值始终等于累计净值时，才可使用 `unit`。
- 只有固定差值已在历史样本中得到证明，并明确设置 `cumulative_offset` 时，才可使用 `offset`。
- 使用 `series_start`，防止观察期、模拟期、买入前历史或接管前旧区间进入新的持仓序列。默认 `--history-scope tail` 使用原表最后日期的下一天并设置 `baseline_overlap: last_existing_point`；此组合表示“写入边界而非公式重置”：只允许从 `series_start` 写入候选净值，最后既有点用于只读核验，但新增首日收益、累计汇总、基准和超额公式仍延续原表完整历史。`--history-scope mail-history` 不生成这两个字段，使邮箱历史进入受监督预览；没有 `baseline_overlap` 的普通 `series_start` 才表示真正的序列重置，不连接边界前公式。两种范围都不授权静默回插或修改历史。
- `return_frequency` 必须明确选择 `daily` 或 `weekly`，它决定基准与累计行采用的主收益列。内置模板要求它与 `data_frequency` 相同。
- `data_frequency` 控制实际保留的数据行，默认 `auto`：优先从工作簿既有日期规律判断；历史不足时再从“周收益/日收益”等明确表头判断；仍无证据的空白追加表才沿用邮件本身的记录频率。只有历史和表头都无法证明时，AI 才能依据本地业务证据显式设为 `daily` 或 `weekly`。显式设置与现有模板冲突会被程序拒绝。
- `max_staleness_days` 应等于产品实际披露周期，再加上审慎的节假日缓冲。数据源过期时应阻止预览，而不是把“没有新数据”当作成功。
- 自动补录只会追加晚于工作簿最新净值的日期。若历史中间缺少日期，程序会停止写入；AI 应先对照邮件、备份和工作簿定位原因并提出修复，只有来源相互冲突、无法证明哪个值正确时才询问用户，避免静默重排其他行或跨表公式。

### 历史身份与累计序列检查

`validation-report.json` 每条路由的 `history_integrity` 包含：

- `scope`：`series_start`、边界前保留行数和受管行数；
- `pre_managed_diagnostics`：边界前代码或累计异常的脱敏诊断，只保留、不修改、不计入 `history_repairs_required`；
- `code_column`：检查历史区所有非空产品代码是否恒定，并与路由主代码一致；Excel 自动填充产生的递增伪代码会列入 `unexpected_rows`，只保存行号、日期和代码短指纹。原本空白的重复代码单元格不要求补写。
- `cumulative_sequence`：检查单位净值缺失、累计净值缺失、明确 `unit/offset` 策略违背，以及 `require` 策略中“前后单位/累计价差一致但中间单行跳变”的孤立断点。持续的新价差可能来自真实分红，不会只凭形态自动判错。
- `repair_required`：任一结构检查失败即为真；顶层 `history_repairs_required` 汇总受影响路由数。

每条路由报告另含 `boundary_anchor_verified` / `boundary_anchor_date`，说明 `baseline_overlap` 的最后既有点是否与邮箱唯一一致；它可批准零新增基线，但不参与写入或受管历史修复。续接模式的新增首日收益和汇总公式从原表完整历史取前点与区间，不把这个核验锚点误当成公式历史起点。

这些检查只负责发现，不授权修复。Agent 应用正式净值通知核对受影响日期，并在 `previews/` 中生成只改已证明错误单元格的修复副本和脱敏差异清单；其他历史行、公式、样式和 Sheet 必须逐项证明未变。用户像第一次预览一样批准后，修复才能通过备份、同目录临时副本、Excel/WPS COM、验证和原子替换应用。若标准计划还不能表达该修复，先扩展工具再提交，禁止直接编辑正式工作簿。修复后必须重新运行完整 `preview` 并重新批准自动更新。

多路由发现失败后，`route-report.json` 提供两个脱敏定位入口：`route_overlaps` 把同时命中的范围拆成最小路由对，并附现有主题/代码短指纹及应收窄的 `subject_contains`、`parser`、`code` 或 `code_aliases`；`date_conflicts` 给出冲突净值日、候选数量、值短指纹、来源短指纹和受控来源类型。报告不会保存完整主题、附件名或原始净值。Agent 修改配置后，可用下列命令只复查受影响路由：

```powershell
.\.venv\Scripts\python.exe navctl.py discover --sheet "示例产品A" --sheet "示例产品B"
```

`discover --sheet` 是只读路由诊断；未知或暂停 Sheet 会被拒绝。需要同时检查局部工作簿结果时可把命令改成 `preview --sheet`，后者会生成局部预览或基线报告，但 `plan.json` 固定标记 `diagnostic_only: true`、`committable: false`、`final_full_preview_required: true`，`commit` 必须拒绝。两个命令的 `--sheet` 都可重复。局部复查通过后仍要运行一次不带 `--sheet` 的完整 `preview`，检查跨路由组合并生成统一方案；不要因一个局部结果跳过最终完整预览。

当范围内邮件无法解析时，`diagnostics.notice_classification` 区分 `likely-non-nav-notice`、内置非净值类别和 `nav-parser-gap`；`suggested_filter` 提示先核实后增加 `subject_excludes`，或检查解析器。Excel/CSV/PDF 等附件按附件短指纹列入 `attachment_diagnostics`，报告状态为 `failed` 或 `no-nav-records`，并给出资源上限/无结构化净值/无效内容等受控原因。必须先确认附件是不是正式净值来源；已证明是月报、季报或交易材料时应修正主题范围，不要误补 parser。

### 正式更正通知

普通晚到邮件、重复发送或 UID 更大的邮件不能自动覆盖同日不同值。核心运行时只在以下条件同时满足时采用更正值：

1. 原始行和更正行已分别通过发件人、产品代码/别名或唯一主题绑定，落在同一活动路由和同一净值日；
2. 更正邮件主题明确包含“更正”“修正”“以此为准”或 `corrected`、`correction`、`revised` 等更正语义，而不是普通 `resend`/重发；
3. 邮箱 UID 能证明更正通知后发；仅测试或特殊本地来源没有 UID 时，必须有带时区且可比较的邮件日期，所有候选使用同一种顺序依据；
4. 同一封更正邮件对该产品和日期只能解析出一个单位/累计净值组合；多个更正值仍失败关闭；
5. 更正通知必须晚于该日期的其他全部不同值候选；之后又出现普通不同值时仍视为冲突。

满足条件时，`route-report.json` 的 `corrections_applied` 记录路由、日期、受控更正标记类别、顺序依据，以及替代和被替代候选的消息、主题、数值和来源短指纹。它不写完整主题、附件名或净值。Agent 必须在首次预览中核对替代关系；不满足条件的候选继续进入 `date_conflicts`，不能增加“最后一封自动赢”的宽泛本地规则。

## 批量托管邮件的选择式路由与错误反例

托管人可能在同一封正文或附件中提供十几个产品，而用户只管理其中一部分。假设“示例代码A”已经映射到“示例产品A”，“示例代码B”和“示例代码C”没有活动路由：程序只把前者交给受管工作表，后两者按明确代码忽略。`route-report.json` 只记录忽略行数和涉及邮件数，不写入被忽略的真实代码；一封邮件同时含受管行和未配置行时仍可正常通过。

这种放宽只适用于已经成功解析、且行内代码明确但与所有活动路由都不相等的记录。下列情况继续失败关闭：行没有代码且不能由唯一的 `allow_sender_only` 路由证明归属；同一行匹配多个路由；同日净值冲突；适用解析器失败；整个回看窗口只有未配置行、没有任何受管产品数据；或最新受管数据超过 `max_staleness_days`。暂停产品等同于不在活动路由中，不应拖累同封邮件里的活动产品。

错误反例：Agent 看到 “NAV row was not routed” 后，没有先区分明确未配置代码与真正歧义，而是另建一个约百行的 `nav_sync.py`，硬编码发件人、产品代码、Sheet 和列号，再用 openpyxl/COM 直接写表。即使这个脚本能提取当前两只产品，也错误绕过了首次预览或零新增基线批准、配置与工作簿哈希、历史和公式校验、并发锁、备份、临时副本验证及异常保持原表不变。

一种已经验证的处理方式是：读取脱敏路由诊断并制作虚构样例；明确未配置行使用内置选择式路由；通用 HTML/附件格式可完善 `parser: auto`；供应商专有格式可放进本地 `app/parsers/`。Agent 也可以根据现场证据采用不同的诊断顺序、临时工具或解析库，不必强行沿用这一顺序。无论如何，持久方案最终都继续走 `navctl.py preview`、用户首次批准、`commit` 和正式计划任务，临时直接写表原型不得成为生产入口。

## 工作表模式

下列模式都是 AI 和程序内部配置，不是需要普通用户回答的业务问题；模式不会自动互相降级，AI 也不得为了绕过验证错误而擅自切换。

| `sheet_mode` | 适用结构 | 历史要求 | 汇总行 |
| --- | --- | --- | --- |
| `summary`（默认） | 已有历史数据、固定汇总行和受管理公式的复杂表；也支持严格识别的单一预留数据行及待审查冷启动 | `minimum_history_dates` 是无警告核验目标；产品名称或代码任一唯一对应且至少有一个真实邮件日期即可生成带警告的首次预览 | 数据区后必须紧跟可识别汇总行，冷启动也保留该行 |
| `append` | 空白、只有表头，或没有汇总行的简单追加表 | 允许历史不足的冷启动，但验证报告会显示警告 | 不需要；数据区下方不得有页脚内容 |
| `template` | 只能用于 `workbook init-template` 生成的新表，一只产品一个工作表 | 一个真实邮件日期即可生成带警告的首次预览；两个日期写入并核实后恢复严格验证 | 固定保留累计行，程序只管理收益、基准和超额公式 |

`append` 模式至少配置非空的 `code` 或 `product_name`，并要求目标表中存在对应的可写列。产品名称和代码应由 AI 根据候选邮件和工作簿确定；无法唯一确定时才询问用户。空白工作表会在预览中初始化标准表头；目标工作簿和目标工作表仍必须提前存在。首次能可靠落表的最低记录是：

- 产品名称或产品代码；
- 净值日期；
- 单位净值；
- 按已确认 `cumulative_policy` 得到的累计净值（如果工作表包含累计净值列）。

冷启动不等于跳过检查。未来日期、同日冲突、重复日期、异常净值跳变、邮件路由歧义和累计净值口径错误仍会阻止预览。用户必须亲自检查新表头、产品标识、日期和净值后，才能批准第一次写入。复杂汇总、图表或专业指标应放在未托管的独立分析工作表或单独文件中，不要放在 `append` 数据区下面。

`summary` 还有一个程序自动识别的冷启动状态：表头下只有一条预留数据行，该行仅含与路由精确一致的产品名称或代码及一个预留日期，单位/累计净值为空，其下立即是“累计/合计”行。邮件至少含一个真实日期、预留行没有备注或其他业务内容时，程序用最早真实邮件替换预留日期并补全该行，再把其余日期插入汇总行之前。AI 不得把这种结构改成 `append`，不得询问用户删除“累计”行，也不得为凑门槛伪造历史数据。

另一种待审查状态是“工作簿有历史，但邮箱只有新数据或不足两个重叠日期”。只要工作簿中的产品名称或产品代码任一与唯一邮件路由一致，至少有一个真实邮件日期，而且邮件日期只落在工作簿已有日期或尾部新日期，程序就生成首次预览并标记 `cold_start_kind: summary-reviewed-preview`。这一步不要求主题再次包含代码，也不要求用户提供自己不知道的产品代码。Agent 应检查本地路由报告，用户再核对预览中的产品、日期和净值；身份明确冲突、多代码混入、同日不同值、历史内部插入、异常跳变或累计口径错误仍失败关闭。

`template` 只能与顶层 `workbook_mode: bundled-template` 一起使用，且要求非空 `code` 或 `product_name`。路由确认后运行：

```powershell
.\.venv\Scripts\python.exe navctl.py workbook init-template
```

程序根据每条路由的 `data_frequency` 和 `benchmark` 选择周度/日度、无指数/有指数页面，并把重复的 `benchmark.source_sheet` 合并成一个可见来源页放在末尾。产品页统一保留空白说明行、浅蓝表头、收益浅黄、超额亮黄、净值四位小数、收益两位百分比和动态红涨绿跌。目标文件存在、工作表名不合法、共享指数定义冲突或频率仍为 `auto` 时拒绝创建；不得先删除已有文件来绕过保护。

已有工作簿的模板高于邮件频率：邮箱历史用于提供和核验真实净值，表内日期规律、表头、公式和汇总结构决定落表方式。周度模板收到日度邮件时，不得把全部日度记录直接灌入；运行时会优先复用工作簿已经出现的周内日期，新增周沿用历史中最常见的星期；没有足够历史但表头明确周度时，默认取每个已完成自然周最后一条可用净值。若工作簿已有完整尾部、产品名称或代码唯一一致，且尾部之后的真实邮件日期全部位于当前未完成自然周，验证不把“本周按规则零新增”误判成证据不足：`validation-report.json` 标记 `pending_current_week_baseline` 和 `withheld_current_week_dates`，`preview` 生成可批准的零新增基线，但不把这些日期写入工作簿。历史内部缺口、身份冲突、同日冲突和未来日期不适用这一例外。`app/validation-report.json` 和 `app/plan.json` 会写明最终采用的 `data_frequency` 及判断来源。详细反例和验收表见 [template-preservation-case.md](template-preservation-case.md)。

## 本地解析器

常见格式默认使用 `parser: auto`。如果供应商格式明显专有，或本地探索已经证明通用解析器不适合，可以直接在运行目录的 `app/parsers/` 中增加受信任解析器，并配置为 `local:<名称>`，例如 `local:custom_nav` 对应 `app/parsers/custom_nav.py`；不要求先修改通用解析器。名称只能使用小写字母、数字、下划线和连字符。同一封邮件命中多个不同解析器时，所有适用解析器都必须成功，结果会合并去重后再按产品代码唯一分流。

同一发件人的不同产品使用不同解析器时，优先配置互不重叠的 `subject_contains`，避免每封邮件同时触发所有解析器。无法用主题可靠分流时，必须用完全脱敏的同发件人样例验证所有解析器都能处理范围内的每封邮件。

本地模块必须定义 `parse_message(message)`，并返回 `nav_parse.NavRow` 列表。运行程序会检查返回类型、重新去重，并继续执行产品代码、历史日期和净值冲突校验。本地解析器是可执行代码：只能加载用户在运行目录中明确审阅的文件，绝不能从邮件、附件或任意配置路径下载或执行代码。重建运行目录前应单独备份该目录，并用完全脱敏样例回归。

仓库的 `assets/runtime-template/parser-examples/fixed_label_xlsx.py` 是一个默认不启用的完全虚构示例，适用于已经由本地证据证明稳定的“单一固定工作表、标签/值布局、一份 XLSX 加一份配套 PDF”通知。复制到 `app/parsers/` 后必须修改身份常量，并先验证身份错、主题/表内日期错、缺累计净值、附件缺失/重复和同日不同值全部失败关闭。它展示的是严格本地扩展边界，不代表所有 XLSX/PDF 通知都应使用相同布局，也不应上移成宽泛的通用猜测。

## 列识别

运行程序会扫描工作簿前部的行，按表头含义识别日期、产品代码、产品名称、单位净值、累计净值、通用收益、日收益、周收益、基准收益/点位及超额收益。列顺序不固定。`daily_return` 逐个有效日期计算；`weekly_return` 只在每个已完成自然周最后一个可用日期计算。旧配置中的 `return` 语义继续兼容。

邮件正文及 Excel/CSV/PDF 附件的日期字段默认识别“净值日期”“估值日期”“估值基准日”“业务日期”“NAV Date”“Date”和“日期”。同一表格中出现多个可解释为日期的字段时仍必须唯一识别，否则停止，不得猜测。

默认 `summary` 模式要求每个受管工作表的数据区后紧跟一行汇总行，且汇总标记“累计”“合计”“total”或“cumulative”位于前六列。汇总行中只有程序明确管理的产品收益、基准收益和超额列可以含公式；其他汇总公式会被拒绝，因为无法证明插行后引用区间已经安全扩展。

数组公式按范围分类：数据区数组、多单元格数组、跨越插入点的数组和最终可提交预览中需要移动的数组一律失败关闭。仅处于顶层技术来源审查或 `benchmark.review_only` 许可审查的不可提交副本允许一个窄例外：数组范围只有自身一个单元格，且该单元格位于受管汇总行的基准点位、基准收益或超额列；程序在插行前清空副本中的公式，在新汇总行记录 `review_array_formulas_cleared`，正式工作簿不变。解决基准来源和许可后，这个例外自动失效；Agent 必须先在受保护副本中用 Excel/WPS COM 或可证明等价的方法重建公式，核对公式作用域和计算缓存值，再生成最终完整预览。不得把审查副本直接提交。

显式 `append` 模式支持空白、只有表头或没有汇总行的简单数据区。空白表默认使用“净值日期、产品代码/产品名称、单位净值、累计单位净值”等标准表头；如果使用 `column_overrides` 自定义空白表布局，必须明确给出日期、单位净值和至少一个产品标识列。已有简单表的数据区下方不得存在页脚、说明或汇总内容，否则程序会停止，避免覆盖未知结构。

表头仍有歧义时，可以配置从 1 开始的列号或 Excel 列字母：

```json
{
  "column_overrides": {
    "Demo Fund": {
      "header_row": 2,
      "date": "A",
      "return": "B",
      "name": "C",
      "unit": "D",
      "code": "E",
      "cumulative": "F",
      "benchmark_return": "H",
      "excess": "I"
    }
  }
}
```

需要分别映射两类收益时，用 `daily_return` 和 `weekly_return` 代替上例的通用 `return`，并分别指向实际列。

不得用显式列映射强行确认不确定的理解；应停止并在本机检查工作簿。

## 基准映射

只能映射到已经核对过历史日期和数值的工作簿工作表：

```json
{
  "benchmark": {
    "source_sheet": "Demo Benchmark",
    "source_type": "aligned_return",
    "source_date": "A",
    "source_value": "B",
    "display_name": "示例指数",
    "technical_source_verified": true,
    "license_source": "示例企业数据许可审批记录",
    "license_approved_by": "示例数据负责人",
    "review_only": false
  }
}
```

`display_name` 可选，只能在资料已核实时填写，用于把内置模板中的“基准指数”显示为“中证1000”等名称。仅当来源列已经与产品的日度或周度观察日期对齐时，才使用 `source_type: aligned_return`。日度指数收益列不是周度基准。指数点位应优先使用 `level`；周度模板可在产品页显示点位，日度模板可直接从来源点位计算日收益，不强制增加点位列。缺少必需来源日期时必须阻止正式写入。

四个来源治理字段由 Agent 根据本地证据维护，不让普通用户填写：

- `technical_source_verified`：指数身份、代码、口径、日期字段和历史重叠已经验证；
- `license_source`：企业许可、授权行情商、合同或内部审批的本地引用；不写账号、令牌或合同正文；
- `license_approved_by`：批准角色或本地审批引用，避免记录不必要的个人信息；
- `review_only`：只表示该映射仍不可提交，不是用户选择的运行模式。

新配置一旦使用其中任一字段，就执行成组校验。技术映射完成但许可仍待确认时，设置 `technical_source_verified: true`、`review_only: true`，可保留 `license_source` 记录已审查条款，但不能填写 `license_approved_by`；预览标记 `benchmark-license-unresolved` 并留空新增基准/超额。许可确认后，补齐 `license_source` 和 `license_approved_by`，把 `review_only` 设为 `false`，再重新完整预览。技术未验证不能进入许可审查；许可未批准不能提交。旧版已经上线且不含这些元数据的 `benchmark` 继续兼容，不强制迁移；新接入或重新审查的来源应使用显式字段。

顶层 `benchmark_review_only` 继续只表示“基准身份或技术来源尚未解决”，不能与 `benchmark` 同时使用。它不是长期配置，也不是跳过基准的开关。`products adopt` 发现活跃但来源未证明的基准/超额列时自动使用它；审查预览不能批准为自动更新。如果指数身份最终无法证明，应在副本中与用户确认业务口径后改用真正无基准的产品页。

技术层面能由权威指数定义、工作簿历史日期与数值唯一对应到持续来源时，Agent 应自动完成映射验证，不把内部代码匹配工作交给用户；但仍要单独核对服务条款、机构合同或内部合规批准。工作簿使用自定义基准名称，且本地资料和权威来源都无法证明唯一对应关系时，属于业务身份歧义；技术映射已唯一但条款只覆盖非商业用途、无法证明当前机构用途在许可内时，属于使用许可歧义。两者都只暂停相关页的基准提交，不得猜指数、解释法律边界或拖住其他已核实产品。已授权行情终端/API、企业数据仓库、管理员提供的受控文件和合规本地来源页均可作为上游；运行时只依赖 `benchmark.source_sheet`，不会把某一家公网接口写死。具体两道门禁和接入边界见 [index-data-sources.md](index-data-sources.md)。

Excel 附件解析过程中，openpyxl 可能对不完全支持的扩展、样式或修复内容发出 Python warning。运行时不把原始 warning 文本写入报告，以免带出附件名或本地路径；`route-proposals.json` 和 `route-report.json` 的 `parser_library_warnings` 只记录固定错误码、库名、来源类型、警告类别及计数。出现记录不等于自动解析失败，但 Agent 必须结合脱敏样例和预览核查，不能静默忽略。

本地来源表尚未建立或存在缺日时，AI 按 [index-data-sources.md](index-data-sources.md) 调查公共指数通道。必须先完成来源代码、指数口径、历史重叠和使用条款核验，再把核实后的日期与点位写入本地来源表；当前运行时不会直接调用公网指数接口。

## 派生专业指标

navctl 的职责是把邮件中的原始净值安全地落为可复核序列，不是替用户决定分析口径。夏普比率、最大回撤、年化收益、波动率、卡玛比率等不会被本工具禁止，但应由 AI 或分析代码基于经验证的净值/收益序列另行计算。

AI 应先采用一致、可解释的行业常用口径并在交付中说明：

- 使用单位净值还是累计净值，以及日度、周度或其他收益频率；
- 年化因子和无风险利率的数值、频率与来源；
- 样本起止日期、序列重置点、缺失值和非交易日处理；
- 最大回撤采用净值峰值到谷值的定义，以及是否需要输出回撤区间。

推荐把公式或结果放在 navctl 未托管的独立分析工作表或单独分析文件。固定范围公式未必会随新增行自动扩展，AI 应优先检查并使用合适的整列引用、动态命名范围或 Excel 表格。不得把自定义分析公式塞进受管工作表的汇总行，也不得把“批准写入净值”解释成“批准修改分析模型”。

## 定时任务

```json
{
  "schedule": [
    {"days": ["MON", "TUE", "WED"], "time": "09:30"}
  ]
}
```

时间使用目标电脑的本地时区。首次配置时，AI 应参考 `app/route-proposals.json` 中的近期到达时间给出推荐，然后明确询问用户希望每周一次、工作日每天或哪些星期运行，以及具体时间；只有用户确认后才能写入 `schedule`。可以配置多个时点。Windows 正式部署不得保持空数组，`doctor` 会用 `schedule-config` 和 `schedule_ready` 显示是否完成。定时任务仅支持 Windows、本地路径和已登录的用户会话。

第一次预览经用户批准并成功执行 `commit --yes-reviewed-preview` 后，当前写表配置会记录到本地 `app/automation-approval.json`。此后计划任务自动执行发现、验证、备份和 COM 写入，不要求逐次预览；内部临时验收副本在运行结束后删除。新增/恢复路由以及工作簿路径、IMAP 范围、列映射、样式或校验规则变化会使批准失效，防止新规则未经第一次验收就自动写入。通过 `products pause` 缩小范围且仍有其他活动产品时会延续原批准；只修改运行时间或保留数量也不会使批准失效。

安装后使用 `navctl.py schedule status` 查看任务、自动更新批准状态、上次/下次运行、写入行数和备份路径。失败详情保存在 `logs/update-YYYYMMDD.log`；失败时正式工作簿保持不变。

状态中的 `last_run_time` 和 `next_run_time` 是 Windows 计划任务的本地墙钟时间，并附带 `local_timezone`。不要把 COM 时间对象携带的伪 UTC 偏移当成任务实际按 UTC 调度。路由诊断使用脱敏的邮件短指纹、附件数量、类型和受控异常类型定位问题，不写入任意异常正文、完整主题、附件名或正文。

## 只读预览与正式表并发变化

有新增数据时，预览文件名固定包含 `preview-只读审查-`；零新增时生成的 `.txt` 报告也设置为系统只读。`plan.json` 保存 `preview_display_name`、`preview_read_only`、正式工作簿哈希和逐 Sheet 脱敏结构清单。AI 打开给用户验收的必须是程序返回的预览路径，并核对 Excel/WPS 窗口标题；不得打开正式表冒充预览，也不得解除只读后让用户在预览中修数。需要修改规则时，修改配置或通用/本地解析器后重新生成预览。

正式工作簿可能在审查期间被用户、另一 Agent、Excel/WPS 自动恢复、网盘同步软件或其他进程保存。程序在建立预览基线、生成完成、提交开始和原子替换前比较正式表；发生变化时旧计划立即不可提交，并在 `app/concurrency-report.json` 记录：

- 检测阶段、计划 ID、预期与当前文件哈希；
- Sheet 顺序是否变化、增删/变化 Sheet 数；
- 每张变化 Sheet 的行列增量、非空与公式单元格计数变化、值/公式摘要是否变化；
- 表头与尾部抽样区域最多 20 个变化坐标；不写单元格原值、新值或公式正文。

若只有文件级元数据、计算缓存或应用保存差异，报告标记 `binary_or_metadata_only: true`，但旧计划仍失效：程序不能证明它与用户刚审查的文件完全相同。AI 应先只读查看报告和当前正式表，判断外部保存来源，必要时与最新预览逐项比较，然后以当前正式表重新运行完整 `preview`。不得自动把当前文件视为新批准基线，不得把旧预览行复制回去，不得删除用户或外部进程新增的数据，也不得用备份回滚。新预览通过后再由用户完成首次验收；已批准自动更新遇到同类变化也应停止本轮，重新验证当前基线。

## 零新增首次验收与错误反例

首次配置可能发生在用户已经手工更新完工作簿之后。此时邮箱最新日期与工作簿最新日期一致、待写入数量为零，是正常状态，不代表配置无需批准，也不能要求用户等待下一封邮件才能上线。

正确流程：

1. `preview` 完成邮箱发现、历史数值、产品身份、表结构、日/周频率、公式和基准校验；
2. 没有新增日期时，不复制一份相同工作簿，而在 `previews/` 生成中文 `.txt` 基线验收报告，并保存绑定当前配置、正式表哈希和报告哈希的 `app/plan.json`；
3. 用户检查报告后仍运行 `commit --yes-reviewed-preview`；程序再次核对配置、正式表、报告和 24 小时有效期，返回 `changed: false`、`approved_baseline: true`，只写入自动更新批准，不创建备份、不启动 COM、不修改正式表；
4. 立即安装计划任务。以后无新增是成功的 no-op，有新增才进入备份、临时副本、COM 验证和原子替换。

错误反例：邮箱中“示例产品A”和“示例产品B”的最新日期都已存在于工作簿。Agent 看到零新增后声称“已经跑过 commit，无需上线”，或另写一个硬编码发件人、产品代码、Sheet 和列号的同步脚本，让它有新增就直接覆盖正式表。

该做法错误，因为“本次没有可写行”只说明工作簿无需变化，不能证明自动更新已经获得用户批准、计划任务已经安装或未来写入路径安全。不得：

- 为了获得首次批准而等待或伪造下一条新数据；
- 把零新增退出描述成已完成 `commit`，却没有生成可验证计划和 `automation-approval.json`；
- 新建绕过 `navctl` 的直接写表脚本或计划任务；
- 因为当前表已更新就跳过首次基线验收、任务安装和状态复查。

## 验证与保留

`validation.minimum_history_dates` 不得低于 `2`，它表示达到多少个历史匹配后可以取消冷启动警告，不是首次预览的硬门槛。`summary`、预留行、`append` 和程序生成的 `template` 都可从一个真实邮件日期开始；在达到该数量前持续要求 Agent 和用户检查首次预览。这不放宽未来日期、身份冲突、重复、同日不同值、历史内部插入、异常跳变、路由歧义或累计口径检查。`max_future_days` 与 `max_period_change` 只能依据有记录的产品行为调整，不能为了让失败运行通过而放宽。

`retention.backup_count`、`preview_count` 和 `log_days` 用于限制本地敏感文件。运行程序只会清理自身 `backups/`、`previews/` 和 `logs/` 目录中的文件。

## 本地敏感文件

以下文件仅能存在于本地运行目录，并应被 Git 忽略：`app/config.json`、`app/route-proposal-headers.json`、`app/route-proposal-progress.json`、`app/route-proposals.partial.json`、`app/route-proposals.json`、`app/route-report.json`、`app/validation-report.json`、`app/concurrency-report.json`、`app/plan.json`、`app/automation-approval.json`、`app/scheduled_tasks.json`、`app/last-scheduled-run.json`、`app/run.lock`、本地 `app/parsers/`、根目录中的预览工作簿、`logs/` 和 `backups/`。正式工作簿保留在用户指定的原路径。Windows 密钥保存在当前用户的本地应用数据目录，并使用 DPAPI 加密。目录分层见 [runtime-layout.md](runtime-layout.md)。

`app/run.lock` 在正常运行结束后保留一条 `idle` 诊断记录，不代表仍被锁定。真正的并发保护由操作系统持有；进程崩溃或断电后会自动释放，不需要手动删除文件。

`app/demo-runs/` 只包含 `navctl.py demo` 生成的虚构邮件状态、虚构工作簿和预览，不会复制真实配置或密钥。检查完成后使用 `navctl.py demo remove --run-id <run_id>` 删除指定演练。

程序会在 IMAP 搜索后再次精确检查解析出的 `From` 地址，但这只是路由验证，并非加密级发件人认证。若伪造邮件构成实质风险，应先在邮箱服务商侧强制执行 DKIM/DMARC，或使用专用邮箱规则，再启用此流程。
