<p align="center">
  <img src="https://raw.githubusercontent.com/bryson0083/templexl/main/docs/assets/templexl_logo.png" alt="templexl logo" width="200">
</p>

# Temple**xl**

<p align="center"><a href="README.md">English</a> | <b>繁體中文</b></p>

> 基於 [openpyxl](https://openpyxl.readthedocs.io/) 的 Excel 模板報表引擎——以 Excel 檔作為模板，將資料渲染成報表。

`templexl` 讓你用熟悉的 Excel 設計報表模板（樣式、表格、公式、圖片、模板圖形／浮水印），
再以 `{{變數}}` 與 `#{{資料表}}` 標籤填入資料，輸出成最終報表。 **無需安裝 Excel** 即可使用。

> ⚠️ 開發中（pre-release）。API 在 1.0 之前可能變動。

📖 文件：[使用者指南](docs/USER_GUIDE.md)（API 與範例）｜[架構說明](docs/ARCHITECTURE.md)（貢獻者/Code Review）｜[可執行範例](examples/)

## 安裝

```bash
pip install templexl
```

或使用 [uv](https://docs.astral.sh/uv/)：

```bash
uv add templexl
```

需求：Python 3.12+。

## 快速開始

```python
import pandas as pd
from templexl import render

result = render(
    template="template.xlsx",
    output="output.xlsx",
    data={
        "oper_name": "王小明",
        "date_rng_desc": "2025/01/01 - 2025/01/31",
        "report_df": pd.DataFrame({"姓名": ["Alice", "Bob"], "部門": ["技術部", "業務部"]}),
    },
)

print(result.output_path)   # 'output.xlsx'
print(result.warnings)      # 非致命警告（如未解析的標籤）
```

需要在渲染與存檔之間插入結構操作，或逐工作表填入不同資料時，改用會話式的
`Book`（見下方「會話式 API」）。

## 模板語法

| 語法 | 說明 |
|------|------|
| `{{變數名稱}}` | 以 `data` 中對應的純量值取代 |
| `#{{資料表名稱}}` | 以 pandas DataFrame 展開為表格（含欄位標題列） |
| `#{{資料表名稱 \| noheader}}` | 展開但略過欄位標題列 |

渲染時會保留模板的樣式、合併儲存格與 Excel 表格（Table）範圍，並把位於展開區域
下方的內容與圖片整列往下位移。

展開出來的每一列，樣式一律複製自**標籤所在的那一列**——框線、字型、填色、數字
格式想長什麼樣，就在模板的標籤列上畫好；本套件不會替資料列加上模板裡沒有的樣式。

模板列上的公式會以 Excel **複製語意**平移到每一列（`=B2*2` → `=B3*2`）；但插入點
**下方既有**的公式只是整列位移，參照原文不變。因此表格的彙總請用**結構化參照**
（見〈Excel 表格合計列〉）。

### Excel 表格合計列

標籤綁定的 Excel 表格若有合計列，渲染後合計列緊接最後一筆資料列保留，內容、
樣式與公式原文不變；`table.ref` 涵蓋合計列，篩選範圍則不含合計列（Excel 慣例）。
`[#Totals]` 結構化參照與引用它的具名範圍在輸出檔中仍可解析。

合計列的公式**必須用結構化參照**（也就是 Excel 勾選合計列時的預設寫法），引擎
不改寫合計列公式：

```
✅ =SUBTOTAL(109,表格1[TABLE_SIZE_MB])   隨表格範圍自動涵蓋全部資料列
❌ =SUM(C2:C2)                            寫死 A1 範圍，插列後不會擴張
```

`#{{df}}`（header 模式）重建欄定義時，以**欄名比對**沿用模板欄的合計列設定；
DataFrame 無資料時表格至少保留一列空白資料列。標籤請勿放在合計列——那樣不會與
表格物件綁定，引擎會發出 warning。細節見
[docs/USER_GUIDE.md](docs/USER_GUIDE.md#excel-表格合計列)。

## API

### `render(template, output, data=None, *, with_report=False, escape_formulas=True) -> RenderResult`

| 參數 | 說明 |
|------|------|
| `template` | 模板檔路徑（`.xlsx` 或 `.xlsm`） |
| `output` | 輸出檔路徑；**若已存在會被覆寫**（以原子寫入取代，中途失敗不留半成品） |
| `data` | 渲染資料字典；鍵為標籤名稱，值為純量或 `pandas.DataFrame` |
| `with_report` | 是否產生除錯用渲染報告（記憶體物件，預設關閉） |
| `escape_formulas` | 是否中和 `data` 中以 `=`/`+`/`-`/`@`/tab/CR 開頭的字串值（預設開啟，防公式注入）；模板內的公式不受影響 |

回傳 `RenderResult`：

- `output_path`：已寫出的輸出檔路徑
- `warnings`：非致命警告清單
- `report`：渲染報告（僅 `with_report=True` 時，否則 `None`）；可 `report.write_json(path)` 自行落地

例外皆繼承自 `TemplateError`：`TemplateNotFoundError`、`FileFormatError`、
`TemplateResourceError`、`StructuralEditError`、`RenderError`。

```python
from templexl import TemplateError

try:
    render("t.xlsx", "o.xlsx", data={...})
except TemplateError as e:
    ...
```

## 會話式 API

`render()` 是一次性的「檔案進、檔案出」。當你需要在**渲染與存檔之間**做事時，
改用 `Book`——它把載入、渲染、結構操作、存檔拆成你可掌握的步驟，全程只有
一次 load/save：

```python
from templexl import Book

with Book("template.xlsx") as bk:
    bk.render({"plan_name": "方案A", "detail_df": df})

    sht = bk.sheets[0]                        # 也可用名稱：bk.sheets["報表"]
    sht.ws.merge_cells("A9:C9")               # openpyxl 原生 API（逃生門）
    sht.delete_columns(11)                    # Excel 等效刪整欄

    bk.save("output.xlsx")                    # 落地必須顯式；離開 with 不會自動存檔
```

逐工作表填入不同資料（例如「複製同一份模板工作表 N 份」）：

```python
with Book("template.xlsx") as bk:
    sample = bk.workbook["Sample"]
    for vehicle, detail_df in data_by_vehicle.items():
        copied = bk.workbook.copy_worksheet(sample)   # openpyxl 原生
        copied.title = vehicle
        bk.sheets[vehicle].render({"vehicle_desc": vehicle, "detail_df": detail_df})
    del bk.workbook["Sample"]
    bk.save("output.xlsx")
```

**設計邊界**：`Book` 只包 openpyxl「做不到或做不對」的事（渲染、Excel 等效刪欄）。
合併儲存格、寫值、複製/刪除工作表等 openpyxl 原生能正確完成的操作，一律經
`Sheet.ws` / `Book.workbook` 逃生門直接使用原生 API——本套件不做 xlwings 模擬層。

`render()` 本身即建於 `Book` 之上，兩者共用同一份渲染編排。

### `Sheet.delete_columns(idx)`

補齊 openpyxl `delete_cols` 缺漏的部分，使結果與 Excel「刪除整欄」等效：
合併範圍縮併/丟棄、欄寬左移、列印範圍收縮。

刪除多欄請**由大到小**逐次呼叫，讓座標全程以原始欄位計算：

```python
for col in sorted([13, 12, 11], reverse=True):
    sht.delete_columns(col)
```

工作表存在下列任一特徵時**拒絕執行**並拋出 `StructuralEditError`（訊息含種類與
位置），不產出靜默損壞的檔案：公式、圖片、圖表、Excel 表格、條件格式、資料驗證。
公式一律拒絕的原因見「已知限制」。

## 安全注意事項

- **公式注入防護（預設開啟）**：`data` 帶入、以 `=`/`+`/`-`/`@`/tab/CR 開頭的字串值
  一律以純文字寫入（值不被改寫），Excel 開啟時不會當作公式執行。若確實需要以資料
  動態注入公式，顯式傳 `escape_formulas=False`——僅在資料來源完全可信時使用。
- **不受信任的模板**：載入前會檢查 zip 解壓總量（上限 2GB）與整體壓縮比（上限 200:1），
  超限拋出 `TemplateResourceError`，防解壓炸彈式資源耗盡。一般業務模板遠低於此上限。
- **輸出覆寫**：`output` 指向既有檔案時會被無條件覆寫；寫入為原子操作
  （暫存檔 + rename），渲染中途失敗不會留下半寫入的檔案。
- **日誌**：診斷日誌（含 DEBUG 等級）僅記錄流程與資料形狀，不輸出儲存格資料值。

## 資料量與記憶體

採 openpyxl 全載記憶體模式（模板樣式保留的必要代價），耗時與峰值記憶體皆隨列數
近似線性成長。參考實測（Apple Silicon、Python 3.12、單表 5 欄）：

| 列數 | 耗時 | 峰值記憶體（RSS） |
|------|------|-------------------|
| 10,000 | ~1.5s | ~0.13GB |
| 100,000 | ~15s | ~0.65GB |
| 500,000 | ~77s | ~2.8GB |

建議：50 萬列以上請確保機器有充足 RAM（500k 列建議 8GB 以上可用記憶體）；
百萬列等級請先以 `tests/benchmark.py` 在目標環境實測，或將資料分批為多個輸出檔。

## 已知限制

- **插入點下方的公式不平移**：引擎只有「複製語意」（模板列的公式複製到新列時
  平移），不會因插入 N 列而改寫下方既有公式的參照；具名範圍、條件格式、資料
  驗證與列印範圍同理。彙總請改用結構化參照（`=SUBTOTAL(109,表格1[金額])`），
  由 Excel 依表格範圍即時解析。
- **圖片**：僅平移既有圖片的列向錨點（保留 rowOff/colOff），不處理欄向位移。
- **預排表頭**：若在模板中自行排版了表格的表頭列，請使用 `#{{資料表 | noheader}}`
  語法明確宣告，引擎不會嘗試自動偵測使用者預排的表頭。
- **刪欄不支援公式**：`delete_columns()` 遇到工作表含公式一律拒絕。原因是兩種
  語意不同——引擎現有的公式平移是 Excel **複製語意**（公式自己移動，相對參照隨之
  平移、`$` 絕對參照不動），而刪欄需要的是 **參照改寫語意**（公式不動也要改：指向
  刪除欄右側的參照連同絕對參照一併左移，指向被刪欄者變 `#REF!`）。兩者只在窄情境
  碰巧等價，硬套會靜默算錯，因此明確拒絕而非勉強處理。

## 開發

```bash
uv sync          # 建立虛擬環境並安裝相依（含 dev 群組）
uv run pytest    # 執行黃金檔案回歸測試
```

測試採「黃金檔案」策略：以基準 `.xlsx` 逐格比對渲染輸出（值、樣式、合併、表格
範圍、公式、圖片錨點），確保重構不改變行為。

## 授權

[MIT](LICENSE) © Bryson Xue
