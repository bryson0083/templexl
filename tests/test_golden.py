"""
黃金檔案回歸測試。

針對每個情境，以目前程式重新渲染並與 tests/golden/ 的基準逐格比對。 
任何差異即視為渲染輸出改變，須在重構過程中逐項判定為「改對」或「改壞」。 
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures import SCENARIOS
from tests.golden_compare import (
    compare_signatures,
    render_quietly,
    workbook_signature,
)
from tests.render_adapter import do_render

GOLDEN_DIR = Path(__file__).resolve().parent / "golden"


@pytest.mark.parametrize(
    "name,template,data_fn",
    SCENARIOS,
    ids=[s[0] for s in SCENARIOS],
)
def test_matches_golden(name, template, data_fn, tmp_path):
    golden = GOLDEN_DIR / f"{name}.xlsx"
    assert golden.exists(), (
        f"缺少黃金基準 {golden}，請先執行: uv run python -m tests.generate_golden"
    )

    out = tmp_path / f"{name}.xlsx"
    render_quietly(lambda: do_render(template, out, data_fn()))
    assert out.exists(), f"渲染未產生輸出檔: {out}"

    diffs = compare_signatures(
        workbook_signature(golden),
        workbook_signature(out),
    )
    assert not diffs, "輸出與黃金基準不一致:\n" + "\n".join(diffs)
