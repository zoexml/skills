---
name: store-category-role-report
description: Generate enterprise-vs-store category role analysis reports from Excel files, including advantage and disadvantage categories, department-specific thresholds, 3/4/5-level comparisons, detailed workbook output, and editable PPT reports. Use when the user provides enterprise and store category role workbooks and asks for 门店/企业品类角色差异、代表性差异品类、优势劣势品类、阈值、Top3 or PPT analysis reports.
---

# Store Category Role Report

Use this skill to generate the full `企业-门店品类角色` analysis package:

- Detail workbook for representative store advantage categories.
- Editable PPT report for advantage/disadvantage category analysis.
- Department-specific thresholds for `生鲜部 / 食品部 / 非食品部`.
- 3/4/5-level category comparisons.
- Store pages covering `优势比例Top3`, `优势得分差距Top3`, `劣势指数Top3`, and `劣势得分差距Top3`.

## Inputs

Require two Excel files:

- Store role workbook: sheets `5级品类`, `4级品类`, `3级品类`.
- Enterprise role workbook: sheets `5级品类`, `4级品类`, `3级品类`.

Expected columns include:

- Store: `c_store_id`, `c_ccode`, `c_cname`, `c_dept_name`, `c_score`, `c_sale`, `c_sale_index`, `c_class_role_id`.
- Enterprise: `c_ccode`, `c_cname`, `c_dept_name`, `c_score`, `c_sale`, `c_sale_index`, `c_class_role_id`.

## Business Definitions

- Advantage category: store is `目标性品类`, enterprise same-level category is not `目标性品类`.
- Disadvantage category: enterprise is `目标性品类`, enterprise target coverage store count is at least 4, category is not `作废`, and the store same-level category is not `目标性品类`.
- Threshold: minimum `c_score` among `目标性品类`, calculated by `层级 + 部门`; store threshold and enterprise threshold are both shown.
- Advantage ratio: `门店得分 / 门店阈值`; larger is stronger.
- Advantage score gap: `门店得分 - 门店阈值`; larger is stronger.
- Disadvantage index: `门店得分 / 门店阈值`; smaller is weaker.
- Disadvantage score gap: `门店阈值 - 门店得分`; larger is weaker.

## Workflow

1. Confirm the input files exist and have the expected sheet names.
2. Run the bundled script:

```bash
python3 scripts/run_store_category_role_report.py \
  --store-xlsx "/path/to/门店下的品类角色.xlsx" \
  --enterprise-xlsx "/path/to/企业下的品类角色.xlsx" \
  --output-dir "/Users/joet/Downloads/大数据畅销品推荐" \
  --tag "企业15门店12"
```

3. Return the generated workbook and PPT links to the user.
4. If layout feedback comes back, edit the bundled PPT generator script and rerun.

## Output Files

For `--tag 企业15门店12`, outputs are:

- `各门店代表性差异品类_企业15门店12.xlsx`
- `各门店代表性差异品类分析报告_企业15门店12.pptx`

## Notes

- Do not overwrite older reports unless explicitly asked.
- Keep store names mapped from codes using the existing mapping in the PPT generator.
- If a `部门 + 层级` cell has fewer than 3 items, report the available items only; this means the source data has fewer than 3 eligible categories for that combination.
