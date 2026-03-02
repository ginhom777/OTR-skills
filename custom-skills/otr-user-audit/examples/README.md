# Examples

- `OTR-template-new.xlsx`: 新版OTR模板（E/F为两列审查项）。

生成示例：

```bash
python3 scripts/convert_to_otr.py \
  --user /path/用户信息.xlsx \
  --template examples/OTR-template-new.xlsx \
  --out "OTR 用户审核表2026年03月.xlsx"
```
