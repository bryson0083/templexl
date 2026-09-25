"""
渲染呼叫轉接層。

讓黃金測試的呼叫方式與「目前 API」解耦，使測試能跨越後續的套件改名 
（excel_template_renderer → templexl）與 API 重塑（render_template → render）
而無需改動每個測試。重構到對應階段時，只需調整此處。 
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SRC = REPO_ROOT / "src"
for _p in (REPO_ROOT, SRC):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))


def do_render(template_path, output_path, data: dict):
    """以目前可用的 API 渲染模板，輸出至 output_path。"""
    # 1) 新 API（task 5 之後）：render(template, output, data) -> RenderResult
    try:
        from templexl import render  # type: ignore
        return render(template=str(template_path), output=str(output_path), data=data)
    except ImportError:
        pass

    # 2) 已改名套件、仍為舊函式（task 4 之後、task 5 之前）
    try:
        from templexl import render_template  # type: ignore
        return render_template(str(template_path), str(output_path), **data)
    except ImportError:
        pass

    # 3) 目前狀態：原套件名 + 舊函式
    from excel_template_renderer import render_template  # type: ignore
    return render_template(str(template_path), str(output_path), **data)


def formula_adjuster():
    """
    取得目前活躍的公式列向平移函式（characterization 用）。

    現況真實邏輯為 cell_ops.translate_formula（openpyxl Translator
    薄包裝，Excel 複製語意）。回傳 None 表示找不到對應內部函式，
    呼叫端應據此 skip（此時 whole-workbook 黃金測試仍守護真實模板的公式）。
    """
    try:
        from templexl.core.cell_ops import translate_formula
        return lambda formula, orig_row, tgt_row: translate_formula(
            formula, orig_row, 1, tgt_row, 1
        )
    except ImportError:
        return None
