---
name: pc-invoice-organizer
version: 1.0.0
author: Bowen
license: MIT
display_name: 本地电脑发票管家
display_name_en: PC Invoice Organizer
description: 电脑硬盘发票归类整理。扫描用户指定目录（含子目录）内的 PDF、OFD、XML、JPG/PNG 发票，抽取发票号码、开票日期、价税合计、买卖方、项目名称，按《企业会计准则》费用性质归类为差旅费、业务招待费、办公费、固定资产等类型，重命名为「开票年月日-类型-金额」，并生成含开票日期、类型、金额（另含税额、发票号码、抬头校验、重复检测、抵扣标记）的 Excel 汇总台账。当用户要求整理发票、发票归类、发票批量重命名、发票汇总成表、发票台账、报销发票整理、电子发票或数电票归档、清理下载文件夹里的发票时使用。
description_zh: 扫描指定目录下的发票文件，自动抽取发票号码、开票日期、价税合计、购销双方与项目名称，按《企业会计准则》费用性质归类，批量重命名为「开票年月日-类型-金额」，并生成含税额、抬头校验、重复检测与抵扣标记的 Excel 台账。适用于整理发票文件夹、发票批量改名、报销发票归类汇总、数电票归档、发票台账制作。
description_en: Scan a folder of invoices (PDF/OFD/XML/images), extract invoice number, issue date, gross amount, buyer, seller and line items, classify each invoice by Chinese accounting expense category, rename them to YYYYMMDD-Category-Amount, and export an Excel ledger with tax amount, header validation, duplicate detection and VAT deductibility marks. Use when organizing, renaming, summarizing or archiving invoices, building an invoice register, or cleaning up a folder full of receipts.
allowed-tools: Bash, Read, Write
---

# 本地电脑发票管家

把散落在硬盘里的发票整理成「可报销、可归档、可查」的状态：文件按规则改名，同时出一份 Excel 台账。

## 一、四步流程

全程用 `scripts/invoice_tool.py`，不要手写改名脚本。

1. **scan**：扫描目录、抽取要素、生成清单 `manifest.json`。不改动任何文件。
2. **复核**：读清单里的 `flags`，把图片发票、解析失败的票据补齐；把分类错误的改成正确类型。
3. **apply**：预览改名方案，用户确认后加 `--execute` 落盘，同时写回滚 CSV。
4. **report**：生成 Excel 汇总，交付给用户。

## 二、执行前必须确认的三件事

只问缺的，用户已经说过的不要重复问：

1. **目录**：要整理的文件夹绝对路径。用户说"我的发票""下载文件夹"时必须问清，不要猜。
2. **公司抬头与税号**：用于校验抬头是否可报销。用户不愿提供就跳过，并在汇报中注明"未做抬头校验"。
3. **改名方式**：原地改名（默认，发票仍留在原目录）／移动到新目录／复制到新目录。用户对原件有顾虑时优先"复制"。

可选但值得问：本次归集的**归属年度**（用 `--year` 标记跨年票据）、固定资产**资本化标准**（默认 5000 元）。

## 三、命令参考

路径含中文或空格时用双引号包裹。

```powershell
# 1. 扫描（默认递归子目录）
python <skill>/scripts/invoice_tool.py scan --src "D:\发票" `
    --company "上海某某科技有限公司" --tax-id "91310230MA1K9XXXXX" --year 2026

# 图片发票补录后重跑（overrides 格式见 references/extraction-guide.md）
python <skill>/scripts/invoice_tool.py scan --src "D:\发票" --overrides "D:\发票\overrides.json"

# 2. 预览改名方案（不落盘）
python <skill>/scripts/invoice_tool.py apply --manifest "D:\发票\_发票整理归档\manifest.json"

# 3. 确认后执行
python <skill>/scripts/invoice_tool.py apply --manifest "D:\发票\_发票整理归档\manifest.json" --execute
#    移动到新目录： --execute --mode move --out "D:\2026年度发票"
#    复制保留原件： --execute --mode copy --out "D:\整理副本"

# 4. 生成 Excel
python <skill>/scripts/invoice_tool.py report --manifest "D:\发票\_发票整理归档\manifest.json"
```

| 参数 | 说明 |
| --- | --- |
| `--no-recursive` | 只扫当前目录，不进子目录 |
| `--profile` | 长期配置 JSON，含 `company`、`tax_id`、`fixed_asset_threshold`、`category_overrides` |
| `--manifest` | 清单路径，默认 `<目录>\_发票整理归档\manifest.json` |
| `--xlsx` | 汇总表路径，默认归档目录下 `发票汇总_时间.xlsx` |
| `--archive` | 归档目录（清单、回滚记录、汇总表），默认 `<输出目录>\_发票整理归档` |

## 四、命名规则

`开票年月日-类型-金额.扩展名`，例：`20261108-差旅费-1280.00.pdf`

- 日期取**开票日期**，`YYYYMMDD`；不是文件名日期，也不是报销日期。
- 类型取归类结果，与 Excel 的「类型」列完全一致。
- 金额取**价税合计**，两位小数，不带千分位。
- 重名追加 `_2`、`_3`；疑似重复的票据同样带序号，由 Excel 的「重复校验」列提示。
- 负数红冲票写作 `红冲880.00`，避免文件名出现双连字符。
- 字段不全（缺日期或缺金额）的票据**保留原文件名**并标记，不生成「日期待补」这类无效名称。
- 一页一票的多页 PDF 拆成单票文件；**原合并文件保留不动**，汇报时提醒用户核对后自行归档。

## 五、图片与扫描件

无文字层的图片发票是唯一必须由你亲自读票的环节：

1. `scan` 后看 `manifest.skipped`，列出被标记「图片待人工读取」的文件。
2. 用 Read 打开图片，读出六要素：发票号码、开票日期、价税合计、购买方、销售方、项目名称。
3. 写 overrides JSON（键为文件名，字段格式见 `references/extraction-guide.md`）。
4. 带 `--overrides` 重跑 `scan`，再继续 apply / report。

**不要猜**：模糊、缺角、金额看不清时字段留空并在 `note` 说明，宁可交人工复核。

## 六、复核与修正

`manifest.json` 的每张发票都可直接改：`category`（改类型）、`date`、`total`、`seller`、`note`。改完重跑 `apply` 与 `report` 即可，不必重跑 `scan`。

`flags` 常见值：`缺发票号码`、`缺开票日期`、`缺金额`、`负数红冲票`、`抬头与公司不符`、`购方税号不符`、`疑似重复（N张同号）`、`分类待确认`、`低于固定资产标准X元，建议费用化`、`归属年度不符`、`图片待人工读取`。

逐项处理建议查 `references/reimbursement-checklist.md` 第五节对照表。

## 七、Excel 结构

| 工作表 | 内容 |
| --- | --- |
| 发票明细 | 开票日期、类型、金额（价税合计）、不含税金额、税额、发票号码、票种、销售方、购买方、抬头校验、重复校验、可抵扣、项目名称、文件名、原始路径、复核提示 |
| 分类汇总 | 类型、张数、金额、占比、税前扣除口径、科目说明 |
| 月度汇总 | 归属月份、张数、金额 |
| 异常待处理 | 逐条问题与建议动作 |
| 说明 | 生成时间、目录、口径、命名规则、归档要求 |

用户只要求「开票日期、类型、金额」时也不要删列：不含税金额与税额是入账与抵扣的依据，抬头校验与重复校验是报销环节的刚需。

## 八、交付与汇报

- Excel 通过 `present_files` 交付；整理后的发票目录用文字说明路径即可。
- 汇报固定四段：**整理结果**（张数、金额合计、类型分布）、**命名示例**（2-3 个真实文件名）、**异常清单**（几条、分别是什么、建议怎么做）、**文件位置**（发票目录 + 归档目录 + 汇总表）。
- 明确区分"已完成重命名"与"仅生成预览"；没有 `--execute` 就不要声称已改名。

## 九、边界

- 不删除、不覆盖原始文件；改名失败即跳过并记录，不做破坏性补救。
- 不判定发票真伪、不承诺可抵扣或可税前扣除、不代做账务处理与纳税申报；查验只提供渠道与材料清单。
- 整理结果基于票面信息自动抽取，入账前必须与票面原件复核——这句话要写进汇报。
- 目录内超过 300 个文件时，先 `scan` 再分批 `apply`，避免长时间占用目录。

## 十、参考资料

| 文件 | 何时读 |
| --- | --- |
| `references/category-rules.md` | 分类拿不准、类型有争议、要向用户解释科目归属或税前扣除限额时 |
| `references/category-rules.json` | 需要改关键词或优先级时（脚本读取此文件） |
| `references/extraction-guide.md` | 图片发票识别、解析失败排查、OFD/XML 异常时 |
| `references/reimbursement-checklist.md` | 用户问「能不能报销」「怎么处理这些异常」「要附什么材料」时 |
