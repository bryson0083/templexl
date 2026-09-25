"""
主要 API 介面。

``render()`` 是一次性 facade——建於 :class:`templexl.book.Book` 之上的糖衣：
開啟 Book → 全簿渲染 → 原子存檔。渲染編排（載入檢查、管線執行、warnings
收集、原子寫入）只存在 ``book.py`` 一份，兩個入口行為一致由結構保證。

需要在渲染與存檔之間插入結構操作（合併儲存格、刪欄）或逐工作表填入不同
資料時，直接使用 ``Book``。 
"""
import logging
from typing import Any, Dict, Optional

from .book import Book
from .exceptions import RenderError, TemplateError
from .result import RenderResult

logger = logging.getLogger(__name__)


def render(
    template: str,
    output: str,
    data: Optional[Dict[str, Any]] = None,
    *,
    with_report: bool = False,
    escape_formulas: bool = True,
) -> RenderResult:
    """以資料渲染 Excel 模板並寫出報表。

    Args:
        template: 模板檔路徑（``.xlsx`` 或 ``.xlsm``）。載入前會檢查 zip
            資源占用（解壓總量/壓縮比上限），防解壓炸彈。
        output: 輸出檔路徑。若已存在會被覆寫；寫入為原子操作
            （同目錄暫存檔 + rename），中途失敗不留下半寫入的檔案。
        data: 渲染資料；鍵為標籤名稱，值為純量或 pandas DataFrame。
        with_report: 是否產生渲染報告（除錯/維護用），預設關閉。
            啟用時報告以記憶體物件掛在 ``RenderResult.report``，不會自動寫入磁碟。
        escape_formulas: 是否中和 ``data`` 中以 ``=``/``+``/``-``/``@``/tab/CR
            開頭的字串值（強制為文字型別，防公式注入），預設開啟。
            模板檔內由模板作者撰寫的公式不受影響。傳 ``False`` 保留
            「``=`` 開頭字串視為公式」的舊行為。

    Returns:
        RenderResult: 含 ``output_path``、``warnings``，以及（啟用時）``report``。

    Raises:
        TemplateNotFoundError: 模板檔不存在。
        FileFormatError: 不支援的檔案格式或無效的 zip 結構。
        TemplateResourceError: 模板 zip 資源占用超過上限（解壓炸彈防護）。
        RenderError: 渲染過程發生錯誤。

    Examples:
        >>> import pandas as pd
        >>> from templexl import render
        >>> result = render(
        ...     "template.xlsx",
        ...     "output.xlsx",
        ...     data={
        ...         "oper_name": "OPER_NAME",
        ...         "report_df": pd.DataFrame({"姓名": ["Alice", "Bob"]}),
        ...     },
        ... )
        >>> result.output_path
        'output.xlsx'
    """
    template = str(template)
    output = str(output)

    try:
        with Book(template) as book:
            book_result = book.render(
                data, escape_formulas=escape_formulas, with_report=with_report
            )
            book.save(output)

        report = book_result.report
        if report is not None:
            # 報告於渲染當下建立（彼時尚無輸出路徑），落地後補上實際路徑。
            report.output_path = output

        return RenderResult(
            output_path=output, report=report, warnings=book_result.warnings
        )

    except TemplateError:
        # 本套件的已知例外（含 TemplateNotFoundError/FileFormatError/RenderError）原樣往上拋
        raise
    except Exception as e:
        raise RenderError(f"渲染過程發生未預期錯誤: {str(e)}")
