"""
公開 API 契約測試（對應 specs/public-api）。 

驗證 render() 簽章、RenderResult、例外階層、warnings、with_report 與匯出範圍。
"""

from __future__ import annotations

import pytest

import templexl
from templexl import (
    FileFormatError,
    RenderError,
    RenderResult,
    TemplateError,
    TemplateNotFoundError,
    render,
)

from tests.fixtures import TEMPLATE_NON_TABLE, non_table_data
from tests.golden_compare import render_quietly


def _render(out, **kwargs):
    return render_quietly(
        lambda: render(str(TEMPLATE_NON_TABLE), str(out), data=non_table_data(), **kwargs)
    )


def test_render_returns_result(tmp_path):
    out = tmp_path / "o.xlsx"
    result = _render(out)
    assert isinstance(result, RenderResult)
    assert result.output_path == str(out)
    assert out.exists()


def test_default_produces_no_report(tmp_path):
    result = _render(tmp_path / "o.xlsx")
    assert result.report is None


def test_with_report_populates_report(tmp_path):
    result = _render(tmp_path / "o.xlsx", with_report=True)
    assert result.report is not None


def test_warnings_capture_unresolved_tag(tmp_path):
    # data 未提供 city 時，模板中的 {{city}} 應以非致命警告回報
    data = non_table_data()
    del data["city"]
    result = render_quietly(
        lambda: render(str(TEMPLATE_NON_TABLE), str(tmp_path / "o.xlsx"), data=data)
    )
    assert any("city" in w for w in result.warnings)


def test_template_not_found_raises():
    with pytest.raises(TemplateNotFoundError):
        render("definitely_missing.xlsx", "o.xlsx", data={})


def test_unsupported_format_raises(tmp_path):
    bad = tmp_path / "x.txt"
    bad.write_text("not excel")
    with pytest.raises(FileFormatError):
        render(str(bad), str(tmp_path / "o.xlsx"), data={})


def test_exception_hierarchy():
    for exc in (TemplateNotFoundError, FileFormatError, RenderError):
        assert issubclass(exc, TemplateError)


def test_public_surface_is_minimal():
    assert set(templexl.__all__) == {
        # 一次性 facade
        "render",
        "RenderResult",
        # 會話式 API
        "Book",
        "Sheet",
        "BookRenderResult",
        "SheetRenderResult",
        # 例外階層
        "TemplateError",
        "TemplateNotFoundError",
        "FileFormatError",
        "TemplateResourceError",
        "StructuralEditError",
        "RenderError",
    }
    # 內部子套件不應出現在公開匯出
    for internal in ("core", "models"):
        assert internal not in templexl.__all__
