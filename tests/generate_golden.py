"""
產生黃金基準輸出。

以「目前的程式」渲染每個情境，輸出存入 tests/golden/。 
務必在重構「之前」執行一次以鎖定 baseline：

    uv run python -m tests.generate_golden
"""

from __future__ import annotations

from pathlib import Path

from tests.fixtures import SCENARIOS
from tests.golden_compare import render_quietly
from tests.render_adapter import do_render

GOLDEN_DIR = Path(__file__).resolve().parent / "golden"


def main() -> None:
    GOLDEN_DIR.mkdir(parents=True, exist_ok=True)
    for name, template, data_fn in SCENARIOS:
        out = GOLDEN_DIR / f"{name}.xlsx"
        data = data_fn()
        render_quietly(lambda: do_render(template, out, data))
        if out.exists():
            print(f"[OK] 已產生黃金基準: {out}  ({out.stat().st_size} bytes)")
        else:
            print(f"[FAIL] 未能產生: {out}")


if __name__ == "__main__":
    main()
