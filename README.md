<p align="center">
  <img src="https://raw.githubusercontent.com/bryson0083/templexl/main/docs/assets/templexl_logo.png" alt="templexl logo" width="200">
</p>

# Temple**xl**

<p align="center"><b>English</b> | <a href="README.zh-TW.md">中文版</a></p>

> An Excel template reporting engine built on [openpyxl](https://openpyxl.readthedocs.io/) — use an Excel file as the template and render your data into a report.

`templexl` lets you design report templates in the Excel you already know (styles, tables, formulas, images, template shapes / watermarks),
then fill them with data via `{{variable}}` and `#{{dataframe}}` tags to produce the final report. **No Excel installation required.**

> ⚠️ Under development (pre-release). The API may change before 1.0.

📖 Docs (Traditional Chinese): [User Guide](docs/USER_GUIDE.md) (API & examples) | [Architecture](docs/ARCHITECTURE.md) (contributors / code review) | [Runnable examples](examples/)

## Installation

```bash
pip install templexl
```

or with [uv](https://docs.astral.sh/uv/):

```bash
uv add templexl
```

Requirements: Python 3.12+.

## Quick Start

```python
import pandas as pd
from templexl import render

result = render(
    template="template.xlsx",
    output="output.xlsx",
    data={
        "oper_name": "John Doe",
        "date_rng_desc": "2025/01/01 - 2025/01/31",
        "report_df": pd.DataFrame({"Name": ["Alice", "Bob"], "Department": ["Engineering", "Sales"]}),
    },
)

print(result.output_path)   # 'output.xlsx'
print(result.warnings)      # non-fatal warnings (e.g. unresolved tags)
```

When you need to perform structural operations between rendering and saving, or fill each worksheet
with different data, use the session-style `Book` instead (see "Session API" below).

## Template Syntax

| Syntax | Description |
|--------|-------------|
| `{{variable_name}}` | Replaced with the matching scalar value in `data` |
| `#{{dataframe_name}}` | Expanded into a table from a pandas DataFrame (including a header row) |
| `#{{dataframe_name \| noheader}}` | Expanded without the header row |

Rendering preserves the template's styles, merged cells, and Excel Table ranges, and shifts content
and images located below the expanded area down by whole rows.

Every expanded row copies its style from **the row the tag sits on** — whatever borders, fonts, fills,
and number formats you want, draw them on the tag row in the template; this package never adds styles
to data rows that aren't in the template.

Formulas on the template row are shifted to each row with Excel **copy semantics** (`=B2*2` → `=B3*2`);
but formulas **already existing below** the insertion point are only moved down with their rows, and
their references are left unchanged. Therefore, use **structured references** for table aggregates
(see "Excel Table Totals Row").

### Excel Table Totals Row

If the Excel Table bound to a tag has a totals row, after rendering the totals row is kept directly
after the last data row, with its content, styles, and formula text unchanged; `table.ref` covers the
totals row, while the auto-filter range excludes it (Excel convention). `[#Totals]` structured references
and named ranges that refer to it remain resolvable in the output file.

Totals row formulas **must use structured references** (the default Excel writes when you enable the
totals row); the engine does not rewrite totals row formulas:

```
✅ =SUBTOTAL(109,Table1[TABLE_SIZE_MB])   automatically covers all data rows as the table grows
❌ =SUM(C2:C2)                            hard-coded A1 range, won't expand after rows are inserted
```

When `#{{df}}` (header mode) rebuilds the column definitions, totals row settings are carried over from
template columns **matched by column name**; when the DataFrame is empty, the table keeps at least one
blank data row. Do not place a tag on the totals row — it won't be bound to the table object, and the
engine will emit a warning. See
[docs/USER_GUIDE.md](docs/USER_GUIDE.md#excel-表格合計列) (Traditional Chinese) for details.

## API

### `render(template, output, data=None, *, with_report=False, escape_formulas=True) -> RenderResult`

| Parameter | Description |
|-----------|-------------|
| `template` | Path to the template file (`.xlsx` or `.xlsm`) |
| `output` | Output file path; **overwritten if it already exists** (replaced via atomic write, so a mid-way failure leaves no partial file) |
| `data` | Render data dict; keys are tag names, values are scalars or `pandas.DataFrame` |
| `with_report` | Whether to generate a render report for debugging (in-memory object, off by default) |
| `escape_formulas` | Whether to neutralize string values in `data` starting with `=`/`+`/`-`/`@`/tab/CR (on by default, prevents formula injection); formulas in the template are unaffected |

Returns a `RenderResult`:

- `output_path`: path of the written output file
- `warnings`: list of non-fatal warnings
- `report`: render report (only when `with_report=True`, otherwise `None`); persist it yourself with `report.write_json(path)`

All exceptions inherit from `TemplateError`: `TemplateNotFoundError`, `FileFormatError`,
`TemplateResourceError`, `StructuralEditError`, `RenderError`.

```python
from templexl import TemplateError

try:
    render("t.xlsx", "o.xlsx", data={...})
except TemplateError as e:
    ...
```

## Session API

`render()` is a one-shot "file in, file out". When you need to do things **between rendering and saving**,
use `Book` — it splits loading, rendering, structural operations, and saving into steps you control,
with only a single load/save throughout:

```python
from templexl import Book

with Book("template.xlsx") as bk:
    bk.render({"plan_name": "Plan A", "detail_df": df})

    sht = bk.sheets[0]                        # or by name: bk.sheets["Report"]
    sht.ws.merge_cells("A9:C9")               # native openpyxl API (escape hatch)
    sht.delete_columns(11)                    # Excel-equivalent delete entire column

    bk.save("output.xlsx")                    # saving must be explicit; leaving `with` does not auto-save
```

Filling each worksheet with different data (e.g. "copy the same template sheet N times"):

```python
with Book("template.xlsx") as bk:
    sample = bk.workbook["Sample"]
    for vehicle, detail_df in data_by_vehicle.items():
        copied = bk.workbook.copy_worksheet(sample)   # native openpyxl
        copied.title = vehicle
        bk.sheets[vehicle].render({"vehicle_desc": vehicle, "detail_df": detail_df})
    del bk.workbook["Sample"]
    bk.save("output.xlsx")
```

**Design boundary**: `Book` only wraps what openpyxl "can't do, or can't do correctly" (rendering,
Excel-equivalent column deletion). Operations openpyxl already handles correctly — merging cells, writing
values, copying/deleting worksheets — should go directly through the native API via the `Sheet.ws` /
`Book.workbook` escape hatches; this package does not provide an xlwings emulation layer.

`render()` itself is built on top of `Book`, and both share the same render orchestration.

### `Sheet.delete_columns(idx)`

Fills in what openpyxl's `delete_cols` leaves out, so the result is equivalent to Excel's "Delete Entire
Column": merged ranges are shrunk/dropped, column widths shift left, and the print area contracts.

To delete multiple columns, call it **from largest to smallest**, so coordinates are always computed
against the original columns:

```python
for col in sorted([13, 12, 11], reverse=True):
    sht.delete_columns(col)
```

If the worksheet has any of the following, the call is **refused** with a `StructuralEditError` (the
message includes the kind and location), rather than producing a silently corrupted file: formulas,
images, charts, Excel Tables, conditional formatting, data validation. See "Known Limitations" for why
formulas are always refused.

## Security Notes

- **Formula injection protection (on by default)**: string values from `data` starting with
  `=`/`+`/`-`/`@`/tab/CR are always written as plain text (the value is not altered), so Excel won't
  execute them as formulas when opening the file. If you really need to inject formulas dynamically from
  data, pass `escape_formulas=False` explicitly — only when the data source is fully trusted.
- **Untrusted templates**: before loading, the total uncompressed zip size (limit 2GB) and overall
  compression ratio (limit 200:1) are checked; exceeding them raises `TemplateResourceError`, preventing
  zip-bomb-style resource exhaustion. Typical business templates are far below these limits.
- **Output overwrite**: if `output` points to an existing file it is overwritten unconditionally; writes
  are atomic (temp file + rename), so a failure mid-render never leaves a half-written file.
- **Logging**: diagnostic logs (including DEBUG level) only record flow and data shapes, never cell data values.

## Data Volume & Memory

Uses openpyxl's full in-memory mode (the necessary cost of preserving template styles); time and peak
memory grow roughly linearly with row count. Reference measurements (Apple Silicon, Python 3.12, single
table with 5 columns):

| Rows | Time | Peak memory (RSS) |
|------|------|-------------------|
| 10,000 | ~1.5s | ~0.13GB |
| 100,000 | ~15s | ~0.65GB |
| 500,000 | ~77s | ~2.8GB |

Recommendation: for 500k+ rows, make sure the machine has enough RAM (8GB+ available memory recommended
for 500k rows); for million-row workloads, benchmark first in the target environment with
`tests/benchmark.py`, or split the data into multiple output files.

## Known Limitations

- **Formulas below the insertion point are not shifted**: the engine only has "copy semantics" (formulas
  on the template row are shifted when copied to new rows); it does not rewrite references in existing
  formulas below because N rows were inserted. The same applies to named ranges, conditional formatting,
  data validation, and print areas. For aggregates, use structured references instead
  (`=SUBTOTAL(109,Table1[Amount])`), which Excel resolves live against the table range.
- **Images**: only the row anchors of existing images are shifted (rowOff/colOff preserved); column
  shifts are not handled.
- **Pre-laid-out headers**: if you lay out a table's header row yourself in the template, declare it
  explicitly with `#{{dataframe | noheader}}`; the engine does not try to auto-detect user-laid-out headers.
- **Column deletion does not support formulas**: `delete_columns()` always refuses worksheets containing
  formulas. The reason is that the two semantics differ — the engine's existing formula shifting uses Excel
  **copy semantics** (the formula itself moves, relative references shift with it, `$` absolute references
  stay put), whereas deleting columns requires **reference-rewrite semantics** (formulas must change even
  if they don't move: references to the right of the deleted column, including absolute ones, shift left,
  and references to the deleted column become `#REF!`). The two only happen to coincide in narrow cases;
  forcing one onto the other would silently compute wrong results, so the operation is explicitly refused
  rather than handled half-way.

## Development

```bash
uv sync          # create the virtual environment and install dependencies (including the dev group)
uv run pytest    # run the golden-file regression tests
```

Tests use a "golden file" strategy: rendered output is compared cell by cell against baseline `.xlsx`
files (values, styles, merges, table ranges, formulas, image anchors), ensuring refactors don't change behavior.

## License

[MIT](LICENSE) © Bryson Xue
