# templexl 範例

以真實模板示範 `templexl` 的典型應用。模板檔複製自本專案的測試資產，
涵蓋一般儲存格模板與 Excel Table（表格物件）模板兩種型態。

| 檔案 | 說明 |
|------|------|
| `render_report.py` | 基本應用：多工作表營運報表——`{{變數}}` 純量替換、`#{{資料表}}` DataFrame 展開、`warnings` 檢視 |
| `render_excel_table.py` | 進階應用：Excel Table 物件模板——表格範圍自動擴增與同步、`with_report=True` 渲染報告 |
| `render_watermark.py` | 模板圖形與浮水印：文字方塊與頁首圖片兩種浮水印在渲染後仍保留，並示範讀回輸出檔清點驗證（資料沿用 `render_excel_table.py`） |
| `template_non_table.xlsx` | 一般儲存格模板（12 個情境工作表，末兩張為浮水印示範） |
| `template_table.xlsx` | Excel Table 物件模板（9 個情境工作表，末兩張為浮水印示範） |

## 執行

於專案根目錄：

```bash
uv run python examples/render_report.py
uv run python examples/render_excel_table.py
uv run python examples/render_watermark.py
```

輸出寫至 `examples/output/`（此目錄不納入版控）。

## 從已安裝的套件執行

範例只依賴公開 API（`from templexl import render`），若你已 `pip install templexl`，
把模板與範例複製到任何位置皆可直接執行，無需本專案原始碼。 
