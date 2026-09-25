# templexl 測試說明

本目錄是 `templexl` 的測試套件，核心策略為 **黃金檔案測試（golden-file testing）**：
以基準 `.xlsx` 逐格比對渲染輸出，確保重構不會悄悄改變行為。 

---

## 設計理念：為什麼用黃金檔案測試？

渲染結果是一個 `.xlsx`——裡面有幾百個儲存格的值、字型、框線、填色、合併範圍、
表格範圍、公式、圖片位置。這種輸出**無法用手寫 `assert` 描述「正確長相」**。

因此改採黃金檔案策略：

```
   重構前（一次性）
   穩定版程式 ── render(範本, 資料) ──▶ 輸出.xlsx ──存成──▶ tests/golden/  ★基準

   每次跑測試
   現在的程式 ── render(範本, 資料) ──▶ 新輸出.xlsx ─┐
                                                    ├─逐格比對→ 一致✅ / 差異❌
                                tests/golden/ ──────┘
```

差異不代表「錯」——它代表「輸出變了，請你有意識地判斷是改對還改錯」。

---

## 目錄結構

```
tests/
├── README.md                  本文件
├── __init__.py
│
├── test_templates/            渲染的「輸入」範本（含 {{標籤}}）
│   ├── template_non_table.xlsx      10 個工作表，純插列渲染情境
│   ├── template_table.xlsx          7 個工作表，含 Excel Table 物件情境
│   └── demo_table_info_template.xlsx  1 個工作表，含合計列的 Excel Table
│                                      （PD012 真實模板改名；A1:C3 表格＋
│                                       SUBTOTAL 合計列＋具名範圍 total_space）
│
├── golden/                    渲染的「輸出」基準（標籤已替換、表格已展開）
│   ├── non_table.xlsx
│   ├── table.xlsx
│   └── table_totals.xlsx
│
├── fixtures.py                確定性測試資料 + 範本路徑 + 情境清單(SCENARIOS)
├── golden_compare.py          比對引擎：.xlsx → 簽章 → 逐項 diff
├── render_adapter.py          渲染呼叫轉接層（與公開 API 解耦）
├── generate_golden.py         產生/重產黃金基準的腳本（非測試，手動執行）
├── benchmark.py               大量資料效能量測（非測試，手動執行）
│
├── test_golden.py             黃金回歸測試（whole-workbook 逐格比對）
├── test_characterization.py   聚焦案例：鎖定「正確行為」與「已知缺陷」
├── test_table_totals_row.py   Excel 表格幾何契約（合計列、篩選範圍、body ≥ 1）
└── test_public_api.py         公開 API 契約測試（render / RenderResult / 例外）
```

### 三套黃金情境

| 情境 | 範本 | 鎖住什麼 |
|------|------|----------|
| `non_table` | `template_non_table.xlsx` | 10 個工作表的純插列渲染（無 Excel Table 物件） |
| `table` | `template_table.xlsx` | 7 個工作表、多表格與疊放，皆**無合計列** |
| `table_totals` | `demo_table_info_template.xlsx` | 含**合計列**的真實表格：`ref` 涵蓋合計列、篩選範圍不含合計列、`[#Totals]` 具名範圍與 `SUBTOTAL` 公式原文不變、缺失值寫成空白儲存格 |

> 範本（`test_templates/`）與基準（`golden/`）是**輸入**與**輸出**的關係，兩者都需要、
> 且必須分開——不可互相取代。

---

## 快速開始

```bash
uv sync              # 建立虛擬環境並安裝相依（含 dev 群組的 pytest）
uv run pytest        # 執行全部測試（pyproject 已設 testpaths=["tests"]）
```

預期結果：全數通過（目前 124 passed, 1 skipped）。

### 常用指令

| 目的 | 指令 |
|------|------|
| 跑全部測試 | `uv run pytest` |
| 跑單一檔 | `uv run pytest tests/test_golden.py` |
| 依關鍵字篩選 | `uv run pytest -k golden` |
| 顯示詳細輸出 | `uv run pytest -v` |
| 失敗時看完整 diff | `uv run pytest -vv` |

---

## 三支測試各自負責什麼

### 1. `test_golden.py` — 整份輸出回歸
對每個情境重新渲染，與 `golden/` 基準**逐格比對**（值、數字格式、字型/框線/填色/
對齊、合併範圍、`table.ref`/`autoFilter.ref`/合計列旗標/欄定義、全簿具名範圍、
公式字串、圖片錨點）。任何差異即失敗。這是最廣的安全網。

> 簽章結構為 `{"sheets": {工作表名: 簽章}, "defined_names": {...}}`——具名範圍
> 屬全簿層級，故多包一層而非與工作表平放。

### 2. `test_characterization.py` — 聚焦的具名案例
人類可讀地記錄「關鍵行為」，並明確分成兩類標籤：

- ✅ **LOCK-CORRECT**：鎖定目前**正確**的行為，重構不得破壞
  （例：`=SUM(A2:A10)` 平移、絕對參照 `$A$2` 不動、多字母欄 `AA` 平移）
- ⚠️ **LOCK-DEFECT**：鎖定目前**已知錯誤**的行為作為基準
  （例：`LOG10`→`LOG13` 函式名被破壞、跨表參照過度平移、`{{city}}` 未被替換）

> 修補已知缺陷時，請刻意更新對應的 LOCK-DEFECT 斷言——這正是它存在的目的：
> 讓你「確實知道自己改動了它」。

### 3. `test_public_api.py` — 公開 API 契約
驗證 `render()` 簽章、`RenderResult`（`output_path`/`report`/`warnings`）、
例外階層（皆繼承 `TemplateError`）、`with_report` 行為、`__all__` 匯出範圍。

### 4. `test_table_totals_row.py` — Excel 表格幾何契約
以 openpyxl **程式化建構**含合計列的範本（可精準控制 noheader／header／空
DataFrame／疊放／標籤落在合計列／未渲染表格等形狀），驗證 `ref` 涵蓋合計列、
`autoFilter.ref` 不含合計列、合計列旗標與公式原文保留、body ≥ 1。

> 陷阱：表格 `displayName` **不可長得像儲存格參照**（如 `T1`、`C3`），
> 否則 Excel 會判定檔案損毀而拒開；本檔一律以 `Tbl*` 命名。

---

## 測試輸出檔放在哪？

測試渲染出來的 `.xlsx` 一律寫到 **pytest 的暫存目錄**（系統 temp 下），
**不會留在 `tests/` 或專案內**：

```
/var/folders/.../T/pytest-of-<使用者>/pytest-<N>/test_xxx0/...
```

- pytest 自動只保留最近數個批次，較舊的自動清除。
- 跑完 `uv run pytest` 後專案目錄保持乾淨（測試用 `tmp_path` fixture 隔離）。
- 唯一會寫進 `tests/golden/` 的是手動執行的 `generate_golden.py`（見下）。

---

## 重新產生黃金基準

**何時需要：** 當你**有意識地**認可某個輸出變更時（例如修了一個缺陷、或確認某個
效能優化造成的差異是良性改良）。平常不該跑。 

```bash
uv run python -m tests.generate_golden
```

這會以**目前的程式**重新渲染所有情境，覆寫 `tests/golden/` 下的基準檔。

> ⚠️ 重要：重產基準等於「把現在的輸出當成新的正確答案」。請務必**先用
> `uv run pytest` 看清楚差異、確認每一項都是預期的改變**，再重產，否則會把
> bug 一起「祝福」成基準。

---

## 效能基準量測（benchmark）

非自動化測試，用於量測大量資料渲染的耗時與峰值記憶體。每個資料量級各跑一次
乾淨子程序：

```bash
uv run python -m tests.benchmark 10000      # 1 萬列
uv run python -m tests.benchmark 100000     # 10 萬列
uv run python -m tests.benchmark 500000     # 50 萬列
```

輸出 CSV 一行：`rows,seconds,peak_rss_mb`。它會自建一個最小單表範本，餵入 N 列
DataFrame 後渲染。參考值：10k≈1.5s、100k≈15s（近似線性）。

---

## 如何新增一個測試情境

以「新增一個範本 + 資料」為例：

1. 把新範本（含 `{{標籤}}`）放到 `tests/test_templates/`。
2. 在 [tests/fixtures.py](tests/fixtures.py) 新增一個確定性的資料建構函式
   （**不可有隨機值或時間戳記**，否則基準會不穩定），並加入 `SCENARIOS` 清單：
   ```python
   SCENARIOS = [
       ("non_table", TEMPLATE_NON_TABLE, non_table_data),
       ("table", TEMPLATE_TABLE, table_data),
       ("table_totals", TEMPLATE_TABLE_TOTALS, table_totals_data),
       ("my_case", TEMPLATES_DIR / "my_template.xlsx", my_case_data),  # 新增
   ]
   ```
3. 產生它的基準：`uv run python -m tests.generate_golden`

   > ⚠️ 這支腳本會**重產全部**情境的基準。只想新增一套時，跑完務必用
   > `git status tests/golden/` 確認既有基準沒被動到（有動就還原，見上一節）。
4. `uv run pytest` 確認 `test_golden.py` 會自動涵蓋新情境（它對 `SCENARIOS` 參數化）。

若要新增「聚焦行為」斷言，於 [tests/test_characterization.py](tests/test_characterization.py)
依 LOCK-CORRECT / LOCK-DEFECT 分類加上具名測試。

---

## 比對引擎如何運作（進階）

[tests/golden_compare.py](tests/golden_compare.py) 的核心：

- `workbook_signature(path)`：把 `.xlsx` 拆成穩定的巢狀 dict「簽章」——每張工作表
  的每個有值/有樣式儲存格，記錄值、`data_type`、數字格式、字型、填色、對齊、框線；
  再加上合併範圍、Excel Table 的 `ref`/`autoFilter.ref`、圖片錨點列。
- `compare_signatures(golden, actual)`：遞迴 diff 兩份簽章，回傳人類可讀的差異清單。
- `render_quietly(callable)`：執行渲染並抑制過往遺留的除錯輸出（保持測試輸出乾淨）。

[tests/render_adapter.py](tests/render_adapter.py) 提供 `do_render(...)`，讓測試與公開
API 細節解耦（API 演進時只需改這一處）。

---

## 疑難排解

| 症狀 | 可能原因 / 處理 |
|------|----------------|
| `缺少黃金基準 ...` | 尚未產生基準 → 先跑 `uv run python -m tests.generate_golden` |
| `test_golden` 失敗，diff 一堆空格樣式 | 多半是良性樣式差異；用 `-vv` 看 diff，確認後再決定是否重產基準 |
| 找不到 `templexl` 模組 | 先 `uv sync`；測試會自動把 `src/` 加入路徑（見 render_adapter） |
| 基準每次跑都不同 | 檢查 fixtures 的資料是否含隨機/時間值——必須完全確定性 |
