"""
templexl 進階應用範例：Excel Table（表格物件）模板 + 渲染報告。

模板 template_table.xlsx 的 #{{標籤}} 綁定了 Excel「表格」物件
（插入 > 表格，含篩選箭頭與表格樣式）。渲染時 templexl 會：
  - 依資料列數擴增表格範圍，同步 table.ref 與 autoFilter.ref
  - 保留表格樣式與表頭設定（noheader 模板亦支援）
  - 依同工作表前一個表格的實際渲染結果動態推移後續表格位置

本範例同時示範：
  - with_report=True：取得渲染報告（記憶體物件，不落地），
    可檢視每個物件的最終位置與資料形狀
  - escape_formulas 預設防護：資料值以 '=' 開頭時以純文字寫入

執行：
    uv run python examples/render_excel_table.py
輸出：
    examples/output/table_report.xlsx
    examples/output/table_report_render.json（示範由呼叫端自行落地報告）
"""
from pathlib import Path

import pandas as pd

from templexl import TemplateError, render

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "template_table.xlsx"
OUTPUT_DIR = HERE / "output"


def build_data() -> dict:
    staff_df = pd.DataFrame({
        "姓名": ["張三", "李四", "王五"],
        "年齡": [25, 30, 35],
        "部門": ["技術部", "業務部", "人事部"],
        "薪資": [50000, 55000, 60000],
        "補助": [5000, 6000, 7000],
    })
    staff_total_df = staff_df.copy()
    staff_total_df["薪資總額"] = staff_total_df["薪資"] + staff_total_df["補助"]

    vehicle_df = pd.DataFrame({
        "業者": ["保護傘客運", "保護傘客運", "沒有傘客運"],
        "路線": ["A001", "A002", "GA91"],
        "交易筆數": [10, 20, 30],
        "營收金額": [10, 20, 30],
        "平均客單價": [10, 20, 30],
    })

    project_df = pd.DataFrame({
        "專案": ["專案X", "專案Y", "專案Z"],
        "進度": [0.85, 0.60, 0.92],
        "負責人": ["張經理", "李經理", "王經理"],
        "截止日": ["2026/03/31", "2026/04/30", "2026/05/31"],
    })

    category_df = pd.DataFrame({
        "類別": ["類別A", "類別B", "類別C"],
        "計數": [15, 25, 35],
        "比例": [0.2, 0.33, 0.47],
    })
    region_df = pd.DataFrame({
        "區域": ["北區", "中區", "南區", "東區"],
        "業績": [500000, 450000, 480000, 520000],
        "達成率": [1.05, 0.95, 1.01, 1.09],
    })
    finance_df = pd.DataFrame({
        "月份": ["一月", "二月"],
        "收入": [1200000, 1350000],
        "支出": [980000, 1050000],
        "淨利": [220000, 300000],
    })

    customer_df = pd.DataFrame({
        "客戶": ["客戶甲", "客戶乙", "客戶丙", "客戶丁", "客戶戊"],
        "訂單數": [45, 38, 52, 41, 49],
        "訂單金額": [1250000, 980000, 1450000, 1120000, 1350000],
        "回款狀態": ["已回款", "部分回款", "已回款", "未回款", "已回款"],
    })
    warehouse_df = pd.DataFrame({
        "倉庫": ["倉庫1", "倉庫2", "倉庫3", "倉庫4"],
        "庫存量": [5000, 4500, 5200, 4800],
        "周轉率": [12.5, 10.8, 13.2, 11.6],
        "管理員": ["王管理", "李管理", "張管理", "陳管理"],
    })
    department_df = pd.DataFrame({
        "部門": ["行政部", "財務部", "技術部", "業務部", "人資部", "研發部"],
        "人數": [8, 12, 35, 28, 6, 42],
        "平均薪資": [45000, 52000, 68000, 55000, 48000, 72000],
        "預算": [360000, 624000, 2380000, 1540000, 288000, 3024000],
    })

    return {
        "oper_name": "王小明",
        "date_rng_desc": "2026年06月01日-2026年06月30日",
        "plan_name": "TW PASS",
        "ptnr_req_date_rng_desc": "2026/06/01 - 2026/06/30",
        "txn_date_rng_desc": "2026/06/01 - 2026/06/30",
        "rep_vehicle_df": vehicle_df,
        "report_df": staff_df,
        "report2_df": staff_total_df,
        "report3_df": region_df,
        "report4_df": finance_df,
        "scenario5_report2_df": project_df,
        "scenario6_report2_df": category_df,
        "scenario6_report3_df": region_df,
        "scenario6_report4_df": finance_df,
        "scenario7_report2_df": customer_df,
        "scenario7_report3_df": warehouse_df,
        "scenario7_report4_df": department_df,
    }


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    output = OUTPUT_DIR / "table_report.xlsx"

    try:
        result = render(
            str(TEMPLATE),
            str(output),
            data=build_data(),
            with_report=True,  # 渲染報告：除錯/稽核用，預設關閉
        )
    except TemplateError as e:
        raise SystemExit(f"渲染失敗：{e}")

    print(f"已輸出：{result.output_path}")
    if result.warnings:
        print(f"警告 {len(result.warnings)} 則（模板標籤未在 data 中提供值等）")

    # 渲染報告是記憶體物件；是否落地、寫往何處由呼叫端決定 
    report = result.report
    summary = report.summary
    print(f"渲染物件統計：{summary['total_worksheets']} 張工作表、"
          f"{summary['total_objects']} 個物件")
    report_path = OUTPUT_DIR / "table_report_render.json"
    report.write_json(str(report_path))
    print(f"渲染報告已另存：{report_path}")


if __name__ == "__main__":
    main()
