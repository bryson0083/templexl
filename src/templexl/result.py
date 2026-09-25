"""渲染結果物件。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class RenderResult:
    """``render()`` 的回傳結果。

    Attributes:
        output_path: 已寫出的輸出檔路徑。
        report: 可選的渲染報告（僅在 ``with_report=True`` 時產生），預設 ``None``。
        warnings: 渲染過程的非致命警告訊息（例如未解析的標籤）。
    """

    output_path: str
    report: Optional[Any] = None
    warnings: list[str] = field(default_factory=list)


@dataclass
class BookRenderResult:
    """``Book.render()`` 的回傳結果。

    不含 ``output_path``——落地時機由呼叫端經 ``Book.save()`` 決定，
    渲染當下尚無輸出檔路徑可回報。

    Attributes:
        warnings: 全簿渲染的非致命警告訊息（例如未解析的標籤）。
        report: 可選的渲染報告（僅在 ``with_report=True`` 時產生），預設 ``None``。
    """

    warnings: list[str] = field(default_factory=list)
    report: Optional[Any] = None


@dataclass
class SheetRenderResult:
    """``Sheet.render()`` 的回傳結果。

    Attributes:
        sheet_name: 被渲染的工作表名稱。
        warnings: 該工作表的非致命警告訊息（掃描範圍僅限該工作表）。
    """

    sheet_name: str
    warnings: list[str] = field(default_factory=list)
