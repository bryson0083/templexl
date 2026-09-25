"""
黃金測試的確定性 fixtures。

資料內容搬自重構前的示範腳本 test_unified_scenarios.py /
test_unified_scenarios_table_obj.py，刻意保持完全確定性（無隨機、無時間戳記），
使其可作為穩定的黃金基準輸入。 
"""

from pathlib import Path

import pandas as pd

# 測試用範本檔（渲染的「輸入」，含 {{標籤}}）放在 tests/test_templates/ 下。
TESTS_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = TESTS_DIR / "test_templates"
TEMPLATE_TABLE = TEMPLATES_DIR / "template_table.xlsx"
TEMPLATE_NON_TABLE = TEMPLATES_DIR / "template_non_table.xlsx"
TEMPLATE_TABLE_TOTALS = TEMPLATES_DIR / "demo_table_info_template.xlsx"


def non_table_data() -> dict:
    """template_non_table.xlsx 的渲染資料（情境 1-10）。"""
    scenario_1_df = pd.DataFrame({
        '姓名': ['張三', '李四', '王五'],
        '年齡': [25, 30, 35],
        '部門': ['技術部', '業務部', '人事部'],
        '薪資': [50000, 55000, 60000],
        '補助': [5000, 6000, 7000],
    })
    scenario_1_df['薪資總額'] = scenario_1_df['薪資'] + scenario_1_df['補助']

    scenario_5_stat_df = pd.DataFrame({"人數": [45, 38, 52, 23, 8]})

    scenario_6_rep_agg_df = pd.DataFrame({
        "路線": ["紅線", "藍線", "綠線"],
        "浣熊市市民定期票旅次數": [1200, 980, 1150],
        "浣熊市市民定期票補助金額": [24000, 19600, 23000],
        "浣熊市非市民定期票旅次數": [300, 250, 320],
        "浣熊市非市民定期票補助金額": [6000, 5000, 6400],
        "貓熊市民定期票旅次數": [150, 120, 180],
        "貓熊市民定期票補助金額": [3000, 2400, 3600],
        "貓熊非市民定期票旅次數": [80, 60, 90],
        "貓熊非市民定期票補助金額": [1600, 1200, 1800],
        "總補助金額": [34600, 28200, 34800],
    })

    scenario_7_rep_vehicle_df = pd.DataFrame({
        "業者": ["A公司", "B公司", "C公司"],
        "行政區劃": ["中央區", "東區", "西區"],
        "營收款_比例": [0.3, 0.25, 0.45],
        "營收款_筆數": [1200, 980, 1350],
        "營收款_金額": [1500000, 1200000, 1800000],
        "分配票收金額": [1260000, 1050000, 1890000],
        "差額補貼款金額": [240000, 150000, -90000],
        "中央差額補貼款比例": [0.6, 0.6, 0.0],
        "中央差額補貼款金額": [144000, 90000, 0],
        "地方差額補貼款金額": [0, 0, 0],
    })

    scenario_8_tb4_routeid_df = pd.DataFrame({
        "路線IC代碼": ["R001", "R002", "R003"],
        "路線編號": ["101", "102", "103"],
        "路線名稱": ["市中心線", "工業區線", "住宅區線"],
        "學生刷卡旅次量\nA=IC學生票刷卡量": [2500, 1800, 2200],
        "市民卡(學生卡)\n應收總金額(元)\nB=市民卡營收金額": [125000, 90000, 110000],
        "市民卡(學生卡)\n實際營收金額(元)\nC=學生票營收金額": [100000, 72000, 88000],
        "市民卡(學生卡)基本里程優惠金額(元)D": [12500, 9000, 11000],
        "票價上限補貼金額(元)E": [7500, 5400, 6600],
        "學生卡25折補助金額(元)F\nF=B-C-D-E-G-H-I-J": [5000, 3600, 4400],
        "捷運轉乘優惠補貼金額(元)G": [3000, 2160, 2640],
        "A22-A23區間優惠補貼金額(元)H": [0, 0, 0],
        "國道及台鐵轉乘市區客運優惠金額(元)I": [0, 0, 0],
        "其它補貼(元)J": [0, 0, 0],
    })

    def _disable_df(bus, taxi, mrt, rail, sports):
        return pd.DataFrame({
            "年齡": ["65-74", "75-84", "85+"],
            "公車次數": bus[0], "公車點數": bus[1],
            "愛心計程車次數": taxi[0], "愛心計程車點數": taxi[1],
            "桃園捷運次數": mrt[0], "桃園捷運點數": mrt[1],
            "台鐵次數": rail[0], "台鐵點數": rail[1],
            "運動中心次數": sports, "運動中心點數": [0, 0, 0],
            "活動中心次數": [0, 0, 0], "活動中心點數": [0, 0, 0],
        })

    scenario_9_01 = _disable_df(
        ([1200, 800, 300], [24000, 16000, 6000]),
        ([150, 100, 50], [7500, 5000, 2500]),
        ([300, 200, 80], [6000, 4000, 1600]),
        ([80, 50, 20], [1600, 1000, 400]),
        [40, 25, 10],
    )
    scenario_9_02 = _disable_df(
        ([1150, 750, 280], [23000, 15000, 5600]),
        ([140, 95, 45], [7000, 4750, 2250]),
        ([280, 190, 75], [5600, 3800, 1500]),
        ([75, 45, 18], [1500, 900, 360]),
        [38, 22, 8],
    )
    scenario_9_03 = _disable_df(
        ([1300, 850, 320], [26000, 17000, 6400]),
        ([160, 110, 55], [8000, 5500, 2750]),
        ([320, 210, 85], [6400, 4200, 1700]),
        ([85, 55, 22], [1700, 1100, 440]),
        [42, 28, 12],
    )
    scenario_9_04 = _disable_df(
        ([1250, 820, 310], [25000, 16400, 6200]),
        ([155, 105, 52], [7750, 5250, 2600]),
        ([310, 200, 82], [6200, 4000, 1640]),
        ([82, 52, 21], [1640, 1040, 420]),
        [41, 26, 11],
    )

    scenario_10_partner_ticket_df = pd.DataFrame({
        "業者": ["甲客運", "乙客運", "丙客運"],
        "行政區劃": ["信義區", "大安區", "中山區"],
        "營收款": [2000000, 1800000, 2200000],
        "比例": [0.3, 0.27, 0.33],
        "筆數": [1500, 1350, 1650],
        "分配票收金額": [2040000, 1836000, 2244000],
        "差額補貼款金額": [-40000, -36000, -44000],
    })
    scenario_10_partner_center_df = pd.DataFrame({
        "業者": ["甲客運", "乙客運", "丙客運"],
        "行政區劃": ["信義區", "大安區", "中山區"],
        "中央差額補貼款比例": [0.0, 0.0, 0.0],
        "中央差額補貼款金額": [0, 0, 0],
    })
    scenario_10_partner_local_df = pd.DataFrame({
        "業者": ["甲客運", "乙客運", "丙客運"],
        "行政區劃": ["信義區", "大安區", "中山區"],
        "地方差額補貼款金額": [-40000, -36000, -44000],
    })

    return {
        'oper_name': '統一測試操作員',
        'date_rng_desc': '2025年01月01日-2025年01月31日統一測試',
        'report_df': scenario_1_df,
        'stat_df': scenario_5_stat_df,
        'ptnr_req_date_rng_desc': '2024年5月1日-2024年5月31日',
        'txn_date_rng_desc': '2024年5月1日-2024年5月31日',
        'rep_agg_df': scenario_6_rep_agg_df,
        'plan_name': '城市通行',
        'rep_partnerreq_date': '2024年6月30日',
        'purchase_amt': 5000000,
        'all_cost': 800000,
        'rep_vehicle_df': scenario_7_rep_vehicle_df,
        'date_rng_desc1': '2024年7月',
        'tb4_routeid_df': scenario_8_tb4_routeid_df,
        'rep_year': '2024',
        'rep_area_disable_01_df': scenario_9_01,
        'rep_area_disable_02_df': scenario_9_02,
        'rep_area_disable_03_df': scenario_9_03,
        'rep_area_disable_04_df': scenario_9_04,
        'tpass_purchase_amt': 8000000,
        'tpass_all_cost': 1200000,
        'city': '台北市',
        'partner_ticket_df': scenario_10_partner_ticket_df,
        'partner_center_df': scenario_10_partner_center_df,
        'partner_local_df': scenario_10_partner_local_df,
    }


def table_data() -> dict:
    """template_table.xlsx 的渲染資料（情境 1-7，含 Excel Table 物件）。"""
    scenario_1_df = pd.DataFrame({
        '姓名': ['張三', '李四', '王五'],
        '年齡': [25, 30, 35],
        '部門': ['技術部', '業務部', '人事部'],
        '薪資': [50000, 55000, 60000],
        '補助': [5000, 6000, 7000],
    })

    scenario_2_df = pd.DataFrame({
        '姓名': ['張三', '李四', '王五'],
        '年齡': [25, 30, 35],
        '部門': ['技術部', '業務部', '人事部'],
        '薪資': [50000, 55000, 60000],
        '補助': [5000, 6000, 7000],
    })
    scenario_2_df['薪資總額'] = scenario_2_df['薪資'] + scenario_2_df['補助']

    scenario_3_vehicle_df = pd.DataFrame({
        '業者': ['保護傘客運', '保護傘客運', '沒有傘客運'],
        '路線': ['A001', 'A002', 'GA91'],
        '交易筆數': [10, 20, 30],
        '營收金額': [10, 20, 30],
        '平均客單價': [10, 20, 30],
    })

    scenario_5_report2_df = pd.DataFrame({
        '專案': ['專案X', '專案Y', '專案Z'],
        '進度': [0.85, 0.60, 0.92],
        '負責人': ['張經理', '李經理', '王經理'],
        '截止日': ['2025/03/31', '2025/04/30', '2025/05/31'],
    })

    scenario_6_report2_df = pd.DataFrame({
        '類別': ['類別A', '類別B', '類別C'],
        '計數': [15, 25, 35],
        '比例': [0.2, 0.33, 0.47],
    })
    scenario_6_report3_df = pd.DataFrame({
        '區域': ['北區', '中區', '南區', '東區'],
        '業績': [500000, 450000, 480000, 520000],
        '達成率': [1.05, 0.95, 1.01, 1.09],
    })
    scenario_6_report4_df = pd.DataFrame({
        '月份': ['一月', '二月'],
        '收入': [1200000, 1350000],
        '支出': [980000, 1050000],
        '淨利': [220000, 300000],
    })

    scenario_7_report2_df = pd.DataFrame({
        '客戶': ['客戶甲', '客戶乙', '客戶丙', '客戶丁', '客戶戊'],
        '訂單數': [45, 38, 52, 41, 49],
        '訂單金額': [1250000, 980000, 1450000, 1120000, 1350000],
        '回款狀態': ['已回款', '部分回款', '已回款', '未回款', '已回款'],
    })
    scenario_7_report3_df = pd.DataFrame({
        '倉庫': ['倉庫1', '倉庫2', '倉庫3', '倉庫4'],
        '庫存量': [5000, 4500, 5200, 4800],
        '周轉率': [12.5, 10.8, 13.2, 11.6],
        '管理員': ['王管理', '李管理', '張管理', '陳管理'],
    })
    scenario_7_report4_df = pd.DataFrame({
        '部門': ['行政部', '財務部', '技術部', '業務部', '人資部', '研發部'],
        '人數': [8, 12, 35, 28, 6, 42],
        '平均薪資': [45000, 52000, 68000, 55000, 48000, 72000],
        '預算': [360000, 624000, 2380000, 1540000, 288000, 3024000],
    })

    return {
        'oper_name': '表格物件測試操作員',
        'date_rng_desc': '2025年01月01日-2025年01月31日',
        'plan_name': 'TW PASS',
        'ptnr_req_date_rng_desc': '2025/01/01 - 2025/01/31',
        'txn_date_rng_desc': '2025/01/01 - 2025/01/31',
        'rep_vehicle_df': scenario_3_vehicle_df,
        'report_df': scenario_1_df,
        'report2_df': scenario_2_df,
        'report3_df': scenario_6_report3_df,
        'report4_df': scenario_6_report4_df,
        'scenario5_report2_df': scenario_5_report2_df,
        'scenario6_report2_df': scenario_6_report2_df,
        'scenario6_report3_df': scenario_6_report3_df,
        'scenario6_report4_df': scenario_6_report4_df,
        'scenario7_report2_df': scenario_7_report2_df,
        'scenario7_report3_df': scenario_7_report3_df,
        'scenario7_report4_df': scenario_7_report4_df,
    }


def table_totals_data() -> dict:
    """demo_table_info_template.xlsx 的渲染資料（含合計列的 Excel Table）。

    模板為 `A1:C3` 的表格：表頭列＋`#{{rep_df | noheader}}` 模板列＋合計列
    （`合計` / `=SUBTOTAL(109,表格1[TABLE_SIZE_MB])`），另有具名範圍
    `total_space = 表格1[[#Totals],[TABLE_SIZE_MB]]` 供 H1 計算水位。

    資料刻意含一筆 `TABLE_SIZE_MB=0` 與一筆 `TABLE_ROWS` 為 NULL，
    順帶鎖住缺失值正規化（NULL 須寫成空白儲存格，而非字面的 `<NA>`）。
    """
    rep_df = pd.DataFrame({
        "TABLE_NAME": [
            "ARGUS_TXN_LOG",
            "ARGUS_CARD_MASTER",
            "ARGUS_ROUTE_DIM",
            "ARGUS_STAGING_TMP",
            "ARGUS_AUDIT_TRAIL",
        ],
        "TABLE_ROWS": pd.array(
            [1250000, 480000, None, 0, 96500], dtype="Int64"
        ),
        "TABLE_SIZE_MB": [820.5, 310.25, 12.0, 0, 145.75],
    })

    return {"rep_df": rep_df}


# 黃金情境清單：(名稱, 模板路徑, 資料建構函式)
SCENARIOS = [
    ("non_table", TEMPLATE_NON_TABLE, non_table_data),
    ("table", TEMPLATE_TABLE, table_data),
    ("table_totals", TEMPLATE_TABLE_TOTALS, table_totals_data),
]
