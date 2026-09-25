"""
大量資料渲染基準量測（spike）。

單一子程序量測單一資料量級的耗時與峰值記憶體（RSS），可乾淨取得各量級數據。

用法（每個量級各跑一次乾淨子程序）： 
    uv run python -m tests.benchmark 10000
    uv run python -m tests.benchmark 100000
    uv run python -m tests.benchmark 500000
    uv run python -m tests.benchmark 100000 trace   # 追加 tracemalloc 峰值

輸出 CSV 一行：rows,seconds,peak_rss_mb[,tracemalloc_peak_mb]
peak RSS 含 C 層/序列化緩衝的整體峰值；tracemalloc 僅 Python 物件配置，
兩者對照可區分記憶體回歸的來源層。注意 tracemalloc 對配置密集的渲染
有 2-4 倍時間放大，開啟時的 seconds 不可與歷史時間基線比較。
"""
from __future__ import annotations

import sys
import time
import tracemalloc
from pathlib import Path
from tempfile import TemporaryDirectory

if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes
else:
    import resource

import pandas as pd
from openpyxl import Workbook

from tests.golden_compare import render_quietly
from tests.render_adapter import do_render


def _build_template(path: Path) -> None:
    """建立最小單表模板：A1 放置表格標籤，由引擎展開資料。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws["A1"] = "#{{big_df}}"
    wb.save(str(path))


def _make_df(rows: int) -> pd.DataFrame:
    return pd.DataFrame({
        "id": range(rows),
        "name": ["列資料"] * rows,
        "amount": [12345] * rows,
        "ratio": [0.1234] * rows,
        "note": ["x" * 8] * rows,
    })


class _PROCESS_MEMORY_COUNTERS(ctypes.Structure if sys.platform == "win32" else object):
    """psapi 的 PROCESS_MEMORY_COUNTERS；只在 Windows 下被使用。"""

    if sys.platform == "win32":
        _fields_ = [
            ("cb", wintypes.DWORD),
            ("PageFaultCount", wintypes.DWORD),
            ("PeakWorkingSetSize", ctypes.c_size_t),
            ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
            ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t),
            ("PeakPagefileUsage", ctypes.c_size_t),
        ]


def _peak_rss_mb() -> float:
    """本行程的峰值 RSS（MB）。

    Unix 走 ``resource.getrusage``（macOS 的 ru_maxrss 單位為 bytes、Linux 為 KB）；
    Windows 無 ``resource`` 模組，改以 ctypes 呼叫 psapi ``GetProcessMemoryInfo``
    取 ``PeakWorkingSetSize``（語意等同峰值 RSS），不引入額外相依。
    """
    if sys.platform == "win32":
        counters = _PROCESS_MEMORY_COUNTERS()
        counters.cb = ctypes.sizeof(_PROCESS_MEMORY_COUNTERS)
        handle = ctypes.windll.kernel32.GetCurrentProcess()
        if not ctypes.windll.psapi.GetProcessMemoryInfo(
            wintypes.HANDLE(handle), ctypes.byref(counters), counters.cb
        ):
            raise ctypes.WinError(ctypes.get_last_error())
        return counters.PeakWorkingSetSize / (1024 * 1024)

    maxrss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    divisor = 1024 * 1024 if sys.platform == "darwin" else 1024
    return maxrss / divisor


def main(rows: int, trace: bool = False) -> None:
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        template = tmp_path / "tpl.xlsx"
        output = tmp_path / "out.xlsx"
        _build_template(template)
        df = _make_df(rows)

        if trace:
            tracemalloc.start()
        start = time.perf_counter()
        render_quietly(lambda: do_render(template, output, {"big_df": df}))
        elapsed = time.perf_counter() - start

        line = f"{rows},{elapsed:.2f},{_peak_rss_mb():.1f}"
        if trace:
            _, tm_peak = tracemalloc.get_traced_memory()
            tracemalloc.stop()
            line += f",{tm_peak / (1024 * 1024):.1f}"
        print(line)


if __name__ == "__main__":
    main(
        int(sys.argv[1]) if len(sys.argv) > 1 else 10000,
        trace="trace" in sys.argv[2:],
    )
