"""
可觀測性契約測試（對應 specs/observability）。

驗證 DEBUG 等級日誌不拋例外、不含儲存格資料值。 
"""

from __future__ import annotations

import logging

from templexl import render

from tests.fixtures import TEMPLATE_NON_TABLE, non_table_data

SENTINEL = "PII_SENTINEL_9f3a7c"


def _data_with_sentinel() -> dict:
    """把哨兵字串同時放進簡單標籤值與 DataFrame 儲存格。"""
    data = non_table_data()
    data["oper_name"] = SENTINEL
    df = data["report_df"].copy()
    df.loc[0, "姓名"] = SENTINEL
    data["report_df"] = df
    return data


def test_debug_level_render_does_not_raise(tmp_path, caplog):
    """DEBUG 等級下完整渲染不因日誌呼叫拋 TypeError。"""
    with caplog.at_level(logging.DEBUG, logger="templexl"):
        result = render(
            str(TEMPLATE_NON_TABLE), str(tmp_path / "o.xlsx"), data=non_table_data()
        )
    assert (tmp_path / "o.xlsx").exists()
    assert result.output_path == str(tmp_path / "o.xlsx")


def test_debug_logs_do_not_contain_data_values(tmp_path, caplog):
    """DEBUG 日誌僅含流程與形狀資訊，不得出現儲存格資料值。"""
    with caplog.at_level(logging.DEBUG, logger="templexl"):
        render(
            str(TEMPLATE_NON_TABLE), str(tmp_path / "o.xlsx"), data=_data_with_sentinel()
        )
    assert caplog.records, "DEBUG 等級下應有診斷輸出"
    assert SENTINEL not in caplog.text
