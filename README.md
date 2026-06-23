# nav-email-to-excel

把**邮件里的半结构化数据**（正文表格 / .xls(x) 附件 / PDF 周报）安全、可回退地写进一份**人工维护的 Excel**。

这是一个从真实项目里抽离、脱敏后的参考实现：原项目每周自动从多家托管/券商的净值邮件里提取私募基金净值，写入运营人员手工维护的净值表。代码里所有真实基金名、产品代码、邮箱地址都已替换为占位符（`基金01`、`DEMO04`、`*.example.com`）。

> 📖 **先读方法论**：[METHODOLOGY.md（中文）](METHODOLOGY.md) ｜ [METHODOLOGY.en.md (English)](METHODOLOGY.en.md)
> 工具本身只是方法论的一个落地样例——背后那套「安全写入活的电子表格」的打法才是重点。

---

## 它解决什么

- 输入**半结构化**：每个来源格式不同、还会变 → 按来源分别写解析器，配置驱动。
- 目标**人工维护**：有手填怪癖、随时被改、不能写坏 → 读优先、写前备份、提交前预览。
- 写错**后果真实** → 校验先行：先证明能复现已知正确数据，再相信它写新数据。

## 流水线

```
build_index → fill_index → validate → apply ┌ write → 预览副本 ┐ → notify
  抓取索引     拉指数(可选)   回归校验    编排 └ com_sync(COM) → 主表 ┘  汇总邮件
```
每一步都是独立脚本，可单独运行、单独调试。写入分两段：`write` 先把新行算进**一次性预览副本**（不碰主表），`com_sync` 再用 **Excel COM** 把新行移植进真主表——手工格式 100% 保留、不复活主题色（详见 [CLAUDE.md](CLAUDE.md) §4.1）。

## 快速开始

```bash
cd src
pip install -r requirements.txt

# 1) 配置（复制模板，填上你的邮箱/表名/发件人映射）
cp config.example.json   config.json
cp registry.example.json registry.json

# 2) 造一个脱敏样例表来试跑（可选）
python ../examples/make_sample_workbook.py

# 3) 预览（不动正式表，只写副本）
python build_index.py
python write.py

# 4) 确认无误后正式写入（自动先备份）
python write.py --commit

# 或一条龙（平时定时任务用）
python run_weekly.py            # 预览
python run_weekly.py --commit   # 正式
```

## 配置说明

| 文件 | 作用 |
|---|---|
| `config.json` | 邮箱/IMAP、目标表名、发件人→格式映射、通知设置（**含密钥，不入库**） |
| `registry.json` | 每张表的结构（表头行、数据起始行、代码、收益基准列等，**不入库**） |
| `config.example.json` / `registry.example.json` | 入库的模板，照着填 |

`senders` 里的格式键（`gtht`/`citics`/`csc`/`yiyuan`…）对应 `navlib.py` 与 `phase2.py` 里的解析器；接入新来源 = 写一个解析器 + 在 config 里登记，无需改主流程。

## 关键设计（详见方法论）

- **预览/提交两挡**，提交前自动备份带时间戳的副本。
- **格式跟着文档走**：新行复制上一行样式，不在代码里写死字体颜色——操作者改格式只在 Excel 里改。
- **不信任手填数据**：纯空格视为空、写入前 `strip()` 去残留换行。
- **幂等容错**：某张表名对不上/暂缺 → 跳过并记录，不崩整轮。
- **一次性迁移与定时任务分开**：见 `examples/one_time_migration.example.py`（针对某个具体工作簿的改名/重排版，仅作示例，请勿直接套用）。

## 目录

```
src/                     主流程（脱敏）
examples/                样例表生成器 + 一次性迁移示例
METHODOLOGY.md           方法论（中文）
METHODOLOGY.en.md        Methodology (English)
```

## 免责声明

参考实现，按 MIT 许可「按现状」提供。涉及真实财务数据时，务必先在预览/备份下验证，自担风险。

## License

MIT — 见 [LICENSE](LICENSE)。
