"""
templexl 範例：模板圖形與浮水印保留。

template_table.xlsx 的兩張工作表各示範一種浮水印做法：
  工作表8 —— 旋轉的半透明文字方塊（AutoShape）
  工作表9 —— Excel 頁首圖片（插入 > 頁首及頁尾 > 圖片）

openpyxl 只模型化圖片與圖表，上述兩者在它的 load/save 中會**無聲消失**；
templexl 於載入時原樣留存、存檔後回填，因此輸出檔仍保有浮水印。

本範例渲染後直接讀回輸出檔清點保留結果，示範如何在自己的產製流程中
驗證浮水印沒有掉——只用標準函式庫，不需要安裝 Excel。

渲染資料沿用 render_excel_table.py（同一份模板），避免重複定義。

執行：
    uv run python examples/render_watermark.py
輸出：
    examples/output/watermark_report.xlsx
"""
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from posixpath import dirname, join, normpath

from render_excel_table import build_data

from templexl import TemplateError, render

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "template_table.xlsx"
OUTPUT_DIR = HERE / "output"

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "vml": "urn:schemas-microsoft-com:vml",
}
# 錨點的子元素種類 -> 人類看得懂的名稱
KINDS = {"sp": "形狀", "pic": "圖片", "grpSp": "群組", "cxnSp": "連接線", "graphicFrame": "圖表"}


def _rels(zf, part):
    """part 自己的關聯表：{rId: (Type 末段, zip 內完整路徑)}。"""
    rels_part = join(dirname(part), "_rels", Path(part).name + ".rels")
    if rels_part not in zf.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(zf.read(rels_part)):
        target = rel.get("Target")
        resolved = target.lstrip("/") if target.startswith("/") else normpath(join(dirname(part), target))
        out[rel.get("Id")] = (rel.get("Type").rsplit("/", 1)[1], resolved)
    return out


def inspect_drawings(xlsx_path):
    """回傳 {工作表名稱: (錨點種類計數, 頁首頁尾圖片數)}。"""
    result = {}
    with zipfile.ZipFile(xlsx_path) as zf:
        workbook_rels = _rels(zf, "xl/workbook.xml")
        for sheet in ET.fromstring(zf.read("xl/workbook.xml")).iter(f"{{{NS['main']}}}sheet"):
            part = workbook_rels[sheet.get(f"{{{NS['rel']}}}id")][1]
            sheet_rels = _rels(zf, part)

            counts = {}
            hf_images = 0
            for kind, target in sheet_rels.values():
                if kind == "drawing":
                    for anchor in ET.fromstring(zf.read(target)):
                        for child in anchor:
                            name = KINDS.get(child.tag.rsplit("}", 1)[-1])
                            if name:
                                counts[name] = counts.get(name, 0) + 1
                elif kind == "vmlDrawing":
                    vml = ET.fromstring(zf.read(target))
                    hf_images += sum(
                        1 for shape in vml.iter(f"{{{NS['vml']}}}shape")
                        if shape.find(f"{{{NS['vml']}}}imagedata") is not None
                    )
            if counts or hf_images:
                result[sheet.get("name")] = (counts, hf_images)
    return result


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    output = OUTPUT_DIR / "watermark_report.xlsx"

    try:
        result = render(str(TEMPLATE), str(output), data=build_data())
    except TemplateError as e:
        # 圖形擷取或回填失敗時 templexl 拋 RenderError（繼承 TemplateError），
        # 且不會留下缺圖形的輸出檔——寧可失敗也不產出靜默損壞的報表
        raise SystemExit(f"渲染失敗：{e}")

    print(f"已輸出：{result.output_path}\n")

    before = inspect_drawings(TEMPLATE)
    after = inspect_drawings(output)

    def describe(item) -> str:
        counts, hf_images = item
        parts = [f"{kind} {n}" for kind, n in sorted(counts.items())]
        if hf_images:
            parts.append(f"頁首頁尾圖片 {hf_images}")
        return "、".join(parts) or "（無）"

    print("各工作表的模板圖形（模板 → 輸出）：")
    for name in before:
        print(f"  {name}：{describe(before[name])} → {describe(after.get(name, ({}, 0)))}")

    lost = [name for name in before if before[name] != after.get(name, ({}, 0))]
    print("\n每張工作表的圖形數量與模板一致。" if not lost else f"\n以下工作表的圖形數量與模板不符：{lost}")
    print("開啟輸出檔可見：工作表8 的文字方塊浮水印（隨表格展開下移），"
          "工作表9 的頁首圖片浮水印（整頁模式或列印預覽才看得到）。")


if __name__ == "__main__":
    main()
