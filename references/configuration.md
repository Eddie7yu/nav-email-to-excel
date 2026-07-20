# 运行配置说明

## 目录

- [顶层字段](#顶层字段)
- [路由字段](#路由字段)
- [列识别](#列识别)
- [基准映射](#基准映射)
- [定时任务](#定时任务)
- [本地敏感文件](#本地敏感文件)

## 顶层字段

`config.json` 由本地运行实例生成，绝不能提交到 Git。

| 字段 | 含义 |
| --- | --- |
| `schema_version` | 配置结构版本，当前为 `1` |
| `runtime_id` | 随机的部署实例标识，用于隔离密钥和定时任务 |
| `workbook_path` | 正式 `.xlsx` 或 `.xlsm` 工作簿的绝对路径 |
| `imap` | 服务器、端口、用户、邮箱文件夹、回看天数、邮件大小和数量限制；不得包含密码 |
| `routes` | 已授权的“发件人/产品 → 工作表”映射 |
| `column_overrides` | 可选的逐工作表语义列映射 |
| `style` | 优先保留原样式的收益率格式设置 |
| `schedule` | 可选的 Windows 任务定义 |
| `validation` | 历史样本数量和数值容差 |
| `retention` | 本地备份、预览最大数量及日志保留天数 |

## 路由字段

```json
{
  "sender": "sender@example.invalid",
  "subject_contains": "NAV",
  "sheet": "Demo Fund",
  "code": "DEMO01",
  "parser": "auto",
  "allow_sender_only": false,
  "cumulative_policy": "require",
  "cumulative_offset": null,
  "return_basis": "cumulative",
  "return_frequency": "weekly",
  "series_start": "2026-01-01",
  "max_staleness_days": 14,
  "benchmark": null
}
```

- 产品代码按不区分大小写的方式标准化，但标准化后必须完全相等。
- 产品代码必须保存为带引号的 JSON 字符串。系统拒绝数字类型，避免前导零消失。
- 若获准发件人也会发送非净值邮件，应配置 `subject_contains`。范围内的每一封邮件都必须成功解析；解析失败会阻止预览，不能静默跳过。
- 只有发件人长期专用于单一产品、且邮件中不存在稳定产品代码时，才可设置 `allow_sender_only`。
- 累计净值必须来自邮件时，使用 `cumulative_policy: require`。
- 只有历史样本证明单位净值始终等于累计净值时，才可使用 `unit`。
- 只有固定差值已在历史样本中得到证明，并明确设置 `cumulative_offset` 时，才可使用 `offset`。
- 使用 `series_start`，防止观察期、模拟期或买入前历史进入新的持仓序列。
- 必须明确选择 `daily` 或 `weekly` 收益频率。
- `max_staleness_days` 应等于产品实际披露周期，再加上审慎的节假日缓冲。数据源过期时应阻止预览，而不是把“没有新数据”当作成功。
- 自动补录只会追加晚于工作簿最新净值的日期。若历史中间缺少日期，程序会停止并要求人工修复，避免静默重排其他行或跨表公式。

## 列识别

运行程序会扫描工作簿前部的行，按表头含义识别日期、产品代码、产品名称、单位净值、累计净值、收益、基准收益/点位及超额收益。列顺序不固定。

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

不得用显式列映射强行确认不确定的理解；应停止并在本机检查工作簿。

## 基准映射

只能映射到已经核对过历史日期和数值的工作簿工作表：

```json
{
  "benchmark": {
    "source_sheet": "Demo Benchmark",
    "source_type": "aligned_return",
    "source_date": "A",
    "source_value": "B"
  }
}
```

仅当来源列已经与产品的日度或周度观察日期对齐时，才使用 `source_type: aligned_return`。日度指数收益列不是周度基准。指数点位应优先使用 `level`；运行程序会在匹配的产品周期锚点之间计算收益。缺少必需的来源日期时，必须阻止正式写入。

## 定时任务

```json
{
  "schedule": [
    {"days": ["MON", "TUE", "WED"], "time": "09:30"}
  ]
}
```

时间使用目标电脑的本地时区。定时任务仅支持 Windows，要求本地路径和已登录的用户会话，并且只生成预览；绝不写入正式工作簿或发送邮件。

## 验证与保留

`validation.minimum_history_dates` 不得低于 `2`。`max_future_days` 用于阻止未来日期，`max_period_change` 用于在创建预览前拦截不合理的单位净值跳变。只能依据有记录的产品行为调整这些值，不能为了让失败的运行通过而放宽。

`retention.backup_count`、`preview_count` 和 `log_days` 用于限制本地敏感文件。运行程序只会清理自身 `backups/`、`previews/` 和 `logs/` 目录中的文件。

## 本地敏感文件

以下文件仅能存在于本地运行目录，并应被 Git 忽略：`config.json`、`route-report.json`、`validation-report.json`、`plan.json`、预览工作簿、正式工作簿、`logs/`、`backups/` 和 `scheduled_tasks.json`。Windows 密钥保存在当前用户的本地应用数据目录，并使用 DPAPI 加密。

`demo-runs/` 只包含 `navctl.py demo` 生成的虚构邮件状态、虚构工作簿和预览，不会复制真实配置或密钥。检查完成后使用 `navctl.py demo remove --run-id <run_id>` 删除指定演练。

程序会在 IMAP 搜索后再次精确检查解析出的 `From` 地址，但这只是路由验证，并非加密级发件人认证。若伪造邮件构成实质风险，应先在邮箱服务商侧强制执行 DKIM/DMARC，或使用专用邮箱规则，再启用此流程。
