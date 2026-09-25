"""渲染報告——可選的除錯/維護用結構。

僅在 ``render(..., with_report=True)`` 時產生，以記憶體物件回傳，
由呼叫端決定是否經 :meth:`RenderReport.write_json` 落地。

報告內容直接取自渲染後的容器狀態（單一真相來源），不另行重算物件落點，
以避免「報告與實際輸出漂移」的維護陷阱。
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field


@dataclass
class RenderReport:
    """一次渲染的物件清單與摘要。

    Attributes:
        template_path: 模板檔路徑。
        output_path: 輸出檔路徑。
        worksheets: 以工作表名稱為鍵，記錄各工作表的物件清單與計數。
        summary: 整體統計（工作表數、物件數）。
    """

    template_path: str
    output_path: str
    worksheets: dict = field(default_factory=dict)
    summary: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        """轉為純 dict（便於序列化或檢視）。"""
        return asdict(self)

    def write_json(self, path: str, *, indent: int = 2) -> str:
        """將報告寫成 JSON 檔；落地與否、寫往何處完全由呼叫端決定。 

        Args:
            path: 輸出 JSON 路徑。
            indent: 縮排空白數。

        Returns:
            str: 已寫出的檔案路徑。
        """
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(self.to_dict(), fh, ensure_ascii=False, indent=indent)
        return str(path)
