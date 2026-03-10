# CheckOnline

## Excel Transformation Utility

This repository now includes an enterprise-oriented Python utility to transform workbook data based on configurable rules.

### What it does
- Reads input rows from the `Data` sheet.
- Reads transformation rules from the `Criteria` sheet.
- Produces a standardized output in the `Result` sheet.
- If a `Result` sheet already exists, its column order is used as the output format template.

### Run
```bash
python excel_transformer.py /path/to/workbook.xlsx --output /path/to/output.xlsx
```

If `--output` is omitted, the input workbook is updated in-place.

### Criteria columns
| Column | Required | Description |
|---|---|---|
| `target_column` | Yes | Destination field in Result sheet |
| `rule_order` | No | Rule execution order (ascending) |
| `source_column` | No | Source field, or multiple comma-separated fields |
| `operation` | No | `COPY`, `STATIC`, `CONCAT`, `UPPER`, `LOWER`, `TITLE`, `DATE_FORMAT` |
| `operation_arg` | No | Extra argument (separator for CONCAT, format for DATE_FORMAT, value for STATIC) |
| `default_value` | No | Fallback value if transformed result is empty |
| `required` | No | `true/false`; enforces non-empty output |
| `condition` | No | Row-level boolean expression (example: `Country == 'US' and Score >= 75`) |

### Test
```bash
python -m unittest discover -s tests
```

> Note: the unit test suite auto-skips when `pandas` or `openpyxl` are not installed.
