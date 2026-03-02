---
name: otr-user-audit
description: Convert 用户信息 Excel files into OTR 用户审核表 output. Use when the user asks to map user rows into the OTR 8-column sheet, dedupe by 用户ID, and output with filename `OTR 用户审核表YYYY年MM月.xlsx` without preserving template header styles/formulas.
---

# OTR User Audit

## Overview
Use this skill to generate OTR audit sheets from a 用户信息 source workbook and an OTR template workbook.

## Workflow
1. Read source file (`用户信息`) and extract:
   - `A` -> 用户ID
   - `E+F` -> 中文姓名 (fallback `C+D`)
2. Dedupe by 用户ID by default.
3. Write rows into OTR 8 columns A-H:
   - A 用户ID
   - B 用户姓名
   - C `是`（在职）
   - D `是`（使用OTR）
   - E `是`（岗位匹配关系已调整）
   - F `是`（OTR权限已调整）
   - G 经销商ID
   - H 经销商名称
4. Do not preserve template header styles/formulas; output plain table data.
5. Output file name: `OTR 用户审核表YYYY年MM月.xlsx` unless user overrides.

## Commands
Run converter script:

```bash
python3 scripts/convert_to_otr.py \
  --user /path/用户信息.xlsx \
  --template /path/OTR模板.xlsx
```

Override output and dealer info:

```bash
python3 scripts/convert_to_otr.py \
  --user /path/用户信息.xlsx \
  --template /path/OTR模板.xlsx \
  --out "/path/OTR 用户审核表2026年02月.xlsx" \
  --dealer-id SFSZ10 \
  --dealer-name "佛山中升睿之星汽车销售服务有限公司"
```

Disable dedupe only when explicitly requested:

```bash
python3 scripts/convert_to_otr.py --user a.xlsx --template b.xlsx --no-dedupe
```

## Resources
### scripts/
- `scripts/convert_to_otr.py`: XLSX XML-based converter (no third-party Python deps).
