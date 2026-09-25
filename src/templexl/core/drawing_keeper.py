"""模板圖形（AutoShape／文字方塊／群組／連接線）保留。

**為什麼需要這個模組**：openpyxl 只模型化圖片（``<xdr:pic>``）與圖表
（chart ``<xdr:graphicFrame>``）。``openpyxl/reader/drawings.py`` 的
``find_images()`` 契約就是「只取 charts 與 images」，``<xdr:sp>`` 這類圖形
即使被解析成 ``Shape`` 物件也會就地丟棄；存檔時 ``xl/drawings/`` 完全由
``ws._images`` + ``ws._charts`` 重建，因此圖形在任何 load/save round-trip
都會**無聲消失**（不拋例外、不發 warning）。

實務上報表模板的大小章方框、用印框、註記文字方塊都是這種圖形，掉了不會
被程式察覺，只會產出缺框的報表。

**做法**：載入模板時把「openpyxl 無法表示的錨點」連同其 rels／媒體原樣
留存；``Book.save()`` 寫出後於 zip 層回填，並依表格插列量平移錨點列號。
平移採 **Excel 插入列語意**：插入點之前的錨點不動，之後的才位移。
"""

from __future__ import annotations

import io
import logging
import posixpath
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from lxml import etree
from PIL import Image as PILImage

logger = logging.getLogger(__name__)

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

REL_TYPE_DRAWING = f"{NS_REL}/drawing"
REL_TYPE_IMAGE = f"{NS_REL}/image"
REL_TYPE_VML = f"{NS_REL}/vmlDrawing"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"

# CT_Worksheet 中排在 <drawing> 之後的元素；插入時必須排在這些之前。
_AFTER_DRAWING = (
    "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
    "controls", "webPublishItems", "tableParts", "extLst",
)

# CT_Worksheet 中排在 <legacyDrawingHF> 之後的元素（<legacyDrawing> 之後、
# 這些之前）；插入 <legacyDrawingHF> 時必須排在這些之前。
_AFTER_LEGACY_DRAWING_HF = (
    "drawingHF", "picture", "oleObjects", "controls", "webPublishItems",
    "tableParts", "extLst",
)

_MEDIA_CONTENT_TYPES = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "gif": "image/gif", "bmp": "image/bmp", "tiff": "image/tiff",
    "emf": "image/x-emf", "wmf": "image/x-wmf", "svg": "image/svg+xml",
}


@dataclass
class CapturedDrawing:
    """單一工作表中，openpyxl 無法表示而需原樣保留的繪圖內容。

    同時容納兩類保留內容：錨點（形狀、openpyxl 會丟棄的圖片，``anchors``／
    ``rels_xml``／``media``）與頁首頁尾 VML（``hf_vml``／``hf_vml_rels``／
    ``hf_media``）。兩者互不影響，一個工作表可以只有其中一種、兩種都有，
    或兩種都沒有（此時不會出現在 :func:`capture_unsupported_drawings` 的
    回傳字典中）。
    """

    sheet_name: str
    anchors: List[bytes] = field(default_factory=list)
    rels_xml: Optional[bytes] = None
    media: Dict[str, bytes] = field(default_factory=dict)
    hf_vml: Optional[bytes] = None
    hf_vml_rels: Optional[bytes] = None
    hf_media: Dict[str, bytes] = field(default_factory=dict)


# ======================================================================
# 共用小工具
# ======================================================================


def _resolve(base_part: str, target: str) -> str:
    """把 rels 的相對 Target 正規化為 zip 內的完整路徑。"""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels",
                          posixpath.basename(part) + ".rels")


def _sheet_part_map(names: List[str], read) -> Dict[str, str]:
    """回傳 ``{工作表名稱: zip 內的 sheet part 路徑}``。"""
    if "xl/workbook.xml" not in names:
        return {}
    wb = etree.fromstring(read("xl/workbook.xml"))
    rels_part = _rels_path("xl/workbook.xml")
    id_to_target = {}
    if rels_part in names:
        for rel in etree.fromstring(read(rels_part)):
            id_to_target[rel.get("Id")] = _resolve("xl/workbook.xml", rel.get("Target"))

    mapping = {}
    for sheet in wb.iter(f"{{{NS_MAIN}}}sheet"):
        rid = sheet.get(f"{{{NS_REL}}}id")
        target = id_to_target.get(rid)
        if target:
            mapping[sheet.get("name")] = target
    return mapping


def _sheet_drawing_part(sheet_part: str, names: List[str], read) -> Optional[str]:
    """取得工作表所引用的 drawing part 路徑（沒有則 ``None``）。"""
    rels_part = _rels_path(sheet_part)
    if rels_part not in names:
        return None
    for rel in etree.fromstring(read(rels_part)):
        if rel.get("Type") == REL_TYPE_DRAWING:
            return _resolve(sheet_part, rel.get("Target"))
    return None


# ======================================================================
# 擷取判定：鏡像 openpyxl reader.drawings.find_images()
#
# openpyxl 對繪圖中每個錨點的判定（見 find_images() 與
# SpreadsheetDrawing._blip_rels）：直接子元素為 pic 才會嘗試以 Pillow
# 開啟其內嵌媒體，開不了或格式為 WMF（EMF 對 Pillow 而言也回報 WMF）一律
# 丟棄；直接子元素為 graphicFrame 一律視為圖表交給 openpyxl；grpSp 只在
# anchor.groupShape.pic 存在時取出「群組內那一張圖」（其餘形狀隨之丟失，
# 屬 Non-Goal，本模組維持現況不擷取整個群組）；其餘（sp、cxnSp、不含圖片
# 的 grpSp）openpyxl 完全無法表示，一律擷取。
# ======================================================================


def _find_rel(part: str, rel_id: str, names: List[str], read):
    """在 ``part`` 自己的 rels 中找出指定 ``Id`` 的 Relationship 元素。"""
    rels_part = _rels_path(part)
    if rels_part not in names:
        return None
    for rel in etree.fromstring(read(rels_part)):
        if rel.get("Id") == rel_id:
            return rel
    return None


def _openpyxl_keeps_pic(pic_el, drawing_part: str, names: List[str], read) -> bool:
    """判斷 openpyxl 是否會保留這個 ``<xdr:pic>``（鏡像 find_images 的判定）。

    沒有內嵌關聯、關聯目標不是圖片、目標媒體不存在、Pillow 開不了、或
    格式為 WMF（含 EMF），皆視為「openpyxl 保留不了」——本模組需要擷取。
    """
    blip = pic_el.find(f".//{{{NS_A}}}blip")
    if blip is None:
        return False
    rid = blip.get(f"{{{NS_REL}}}embed")
    if not rid:
        return False
    rel = _find_rel(drawing_part, rid, names, read)
    if rel is None or rel.get("Type") != REL_TYPE_IMAGE:
        return False
    target = _resolve(drawing_part, rel.get("Target", ""))
    if target not in names:
        return False
    try:
        image_format = PILImage.open(io.BytesIO(read(target))).format
    except Exception:
        return False
    return (image_format or "").upper() != "WMF"


def _should_capture_anchor(anchor, drawing_part: str, names: List[str], read) -> bool:
    """單一錨點是否需要本模組擷取（openpyxl 保留不了才擷取）。"""
    if anchor.find(f"{{{NS_XDR}}}graphicFrame") is not None:
        return False  # 圖表交給 openpyxl；非圖表 graphicFrame 屬 Non-Goal

    pic = anchor.find(f"{{{NS_XDR}}}pic")
    if pic is not None:
        return not _openpyxl_keeps_pic(pic, drawing_part, names, read)

    grp = anchor.find(f"{{{NS_XDR}}}grpSp")
    if grp is not None:
        # 群組內含圖片：openpyxl 取出群組內那張圖（其餘形狀隨之丟失），
        # 維持現況不擷取整個群組——含圖片的群組屬 Non-Goal。
        return grp.find(f".//{{{NS_XDR}}}pic") is None

    return True  # sp、cxnSp 等 openpyxl 完全無法表示的錨點


# ======================================================================
# 擷取
# ======================================================================


def _capture_header_footer_vml(item: CapturedDrawing, sheet_part: str,
                               names: List[str], read) -> None:
    """擷取工作表的 ``legacyDrawingHF``（頁首頁尾 VML），寫入 ``item`` 的
    ``hf_*`` 欄位（design 決策 4）：VML 位元組、VML 自己的 rels 與其引用的
    媒體位元組原樣留存，整份以 VML 部件為單位搬移，不解析 VML 內容——
    自然涵蓋奇偶頁與首頁的所有位置代碼。工作表沒有 ``legacyDrawingHF``
    時保持 ``item`` 不變（no-op）。
    """
    sheet_rels_part = _rels_path(sheet_part)
    if sheet_rels_part not in names:
        return
    sheet_root = etree.fromstring(read(sheet_part))
    legacy = sheet_root.find(f"{{{NS_MAIN}}}legacyDrawingHF")
    if legacy is None:
        return
    rid = legacy.get(f"{{{NS_REL}}}id")

    vml_part = None
    for rel in etree.fromstring(read(sheet_rels_part)):
        if rel.get("Id") == rid:
            vml_part = _resolve(sheet_part, rel.get("Target"))
            break
    if not vml_part or vml_part not in names:
        return

    item.hf_vml = read(vml_part)
    vml_rels_part = _rels_path(vml_part)
    if vml_rels_part in names:
        item.hf_vml_rels = read(vml_rels_part)
        for rel in etree.fromstring(item.hf_vml_rels):
            target = _resolve(vml_part, rel.get("Target", ""))
            if target in names:
                item.hf_media[target] = read(target)


def capture_unsupported_drawings(template_path: str) -> Dict[str, CapturedDrawing]:
    """掃描模板，留存 openpyxl 會丟棄的繪圖錨點。

    Args:
        template_path: 模板 ``.xlsx``／``.xlsm`` 路徑。

    Returns:
        Dict[str, CapturedDrawing]: 以工作表名稱為鍵；沒有需保留的內容時
        該表不會出現在結果中。

    Raises:
        Exception: 解析失敗時原樣上拋（design 決策 6：fail fast）。本套件
            MUST NOT 產出缺少模板繪圖內容的輸出檔，呼叫端
            （``Book.__init__``）負責包成 ``RenderError``。
    """
    captured: Dict[str, CapturedDrawing] = {}
    with zipfile.ZipFile(template_path) as zf:
        names = zf.namelist()
        read = zf.read
        for sheet_name, sheet_part in _sheet_part_map(names, read).items():
            item = CapturedDrawing(sheet_name=sheet_name)

            drawing_part = _sheet_drawing_part(sheet_part, names, read)
            if drawing_part and drawing_part in names:
                root = etree.fromstring(read(drawing_part))
                keep = [
                    anchor for anchor in root
                    if _should_capture_anchor(anchor, drawing_part, names, read)
                ]
                if keep:
                    item.anchors = [etree.tostring(a) for a in keep]
                    drawing_rels = _rels_path(drawing_part)
                    if drawing_rels in names:
                        item.rels_xml = read(drawing_rels)
                        for rel in etree.fromstring(item.rels_xml):
                            target = _resolve(drawing_part, rel.get("Target", ""))
                            if target in names:
                                item.media[target] = read(target)

            _capture_header_footer_vml(item, sheet_part, names, read)

            if not item.anchors and item.hf_vml is None:
                continue

            captured[sheet_name] = item
            logger.debug(
                "保留工作表 %r 的 %d 個模板圖形錨點、頁首頁尾 VML=%s",
                sheet_name, len(item.anchors), item.hf_vml is not None,
            )
    return captured


def describe_kept_anchors(item: CapturedDrawing) -> List[Tuple[str, Optional[int], Optional[int]]]:
    """回傳 ``item.anchors`` 的 (特徵種類, from_row, from_col) 清單（0-based
    列欄），供 ``Sheet.delete_columns()`` 的刪欄前置檢查使用（design 決策 8）。

    特徵種類：錨點含 ``pic`` 子元素判定為「圖片」（openpyxl 會丟棄而由本
    套件保留者，如 EMF／WMF）；其餘（``sp``、``cxnSp``、不含圖片的
    ``grpSp``）判定為「形狀」。錨點缺 ``from`` 或列欄資訊時該項回傳
    ``None``。不含頁首頁尾 VML——頁首頁尾圖片不綁儲存格，不阻擋刪欄。
    """
    described: List[Tuple[str, Optional[int], Optional[int]]] = []
    for anchor_xml in item.anchors:
        element = etree.fromstring(anchor_xml)
        kind = "圖片" if element.find(f"{{{NS_XDR}}}pic") is not None else "形狀"
        row = col = None
        frm = element.find(f"{{{NS_XDR}}}from")
        if frm is not None:
            row_el = frm.find(f"{{{NS_XDR}}}row")
            col_el = frm.find(f"{{{NS_XDR}}}col")
            row = int(row_el.text) if row_el is not None else None
            col = int(col_el.text) if col_el is not None else None
        described.append((kind, row, col))
    return described


# ======================================================================
# 回填
# ======================================================================


def _shift_anchor(anchor_xml: bytes, shift) -> bytes:
    """依 Excel 插入列語意平移錨點列號。

    ``ShiftInfo.start_row`` 為 1-based 的起算列，錨點列號為 0-based，
    因此門檻取 ``start_row - 1``；門檻之前的錨點不動（這正是「插入」與
    「複製」語意的差別）。
    """
    if shift is None or getattr(shift, "shift_amount", 0) <= 0:
        return anchor_xml

    element = etree.fromstring(anchor_xml)
    threshold = shift.start_row - 1
    for row_el in element.iter(f"{{{NS_XDR}}}row"):
        try:
            row = int(row_el.text)
        except (TypeError, ValueError):
            continue
        if row >= threshold:
            row_el.text = str(row + shift.shift_amount)
    return etree.tostring(element)


def _next_free_rel_id(rels_root) -> str:
    used = {rel.get("Id") for rel in rels_root}
    index = 1
    while f"rId{index}" in used:
        index += 1
    return f"rId{index}"


def _next_free_media_name(parts: Dict[str, bytes], new_parts: Dict[str, bytes],
                          ext: str) -> str:
    """回傳 ``xl/media/image{n}.{ext}`` 的第一個空號（part 名稱）。

    同時檢查輸出既有部件（``parts``）與本次已放置的部件（``new_parts``），
    避免模板媒體檔名與輸出重新編號後同名時互相覆蓋或指錯圖（design 決策 5）。
    """
    index = 1
    while True:
        name = f"xl/media/image{index}.{ext}"
        if name not in parts and name not in new_parts:
            return name
        index += 1


def _place_media(parts: Dict[str, bytes], new_parts: Dict[str, bytes],
                 data: bytes, ext: str, source_part: str) -> str:
    """把媒體寫成第一個空號的 part 名稱，回傳相對於 ``source_part`` 所在
    目錄的 Target（供寫入 Relationship 的 Target 屬性），並補上該副檔名的
    內容型別 Default（design 決策 5，形狀與頁首頁尾 VML 共用的放置函式）。
    """
    ext = ext.lower()
    name = _next_free_media_name(parts, new_parts, ext)
    new_parts[name] = data
    _ensure_media_content_type(parts, ext)
    return posixpath.relpath(name, posixpath.dirname(source_part))


def _relink_anchors(anchors: List[bytes], item: CapturedDrawing,
                    parts: Dict[str, bytes], new_parts: Dict[str, bytes],
                    target_rels_root, drawing_part: str) -> List[bytes]:
    """把錨點內的關聯 ID 重新指派到 ``target_rels_root``，只為錨點實際
    出現的 r:* 屬性建立關聯；圖片類關聯的媒體透過 :func:`_place_media`
    取空號寫入（design 決策 5）。新建 drawing 與合併進既有 drawing 兩條
    路徑共用——``target_rels_root`` 可以是空白 Relationships（新建）或
    既有 drawing 的 rels（合併）。回傳改寫後的錨點清單。
    """
    if not item.rels_xml:
        return anchors

    source = {rel.get("Id"): rel for rel in etree.fromstring(item.rels_xml)}
    id_map: Dict[str, str] = {}
    result = []
    for anchor_xml in anchors:
        element = etree.fromstring(anchor_xml)
        for node in element.iter():
            for attr, value in list(node.attrib.items()):
                if not attr.startswith(f"{{{NS_REL}}}") or value not in source:
                    continue
                if value not in id_map:
                    rel = source[value]
                    new_id = _next_free_rel_id(target_rels_root)
                    added = etree.SubElement(
                        target_rels_root, f"{{{NS_PKG_REL}}}Relationship"
                    )
                    added.set("Id", new_id)
                    added.set("Type", rel.get("Type"))
                    if rel.get("Type") == REL_TYPE_IMAGE and rel.get("TargetMode") != "External":
                        source_target = _resolve(drawing_part, rel.get("Target", ""))
                        ext = posixpath.splitext(source_target)[1].lstrip(".")
                        added.set("Target", _place_media(
                            parts, new_parts, item.media[source_target],
                            ext, drawing_part,
                        ))
                    else:
                        # 外部連結（TargetMode="External"，如 <a:blip r:link>
                        # 指向檔案系統路徑或 URL）不在 zip 內，media 沒有對應
                        # 項次；原樣複製 Target／TargetMode，取得媒體屬 Non-Goal。
                        added.set("Target", rel.get("Target"))
                        if rel.get("TargetMode"):
                            added.set("TargetMode", rel.get("TargetMode"))
                    id_map[value] = new_id
                node.set(attr, id_map[value])
        result.append(etree.tostring(element))
    return result


def _insert_drawing_element(sheet_root, rel_id: str) -> None:
    """在 sheet XML 的合法位置插入 ``<drawing r:id=.../>``。

    CT_Worksheet 對子元素順序有嚴格規定，``<drawing>`` 必須排在
    legacyDrawing／tableParts／extLst 這些元素之前，否則 Excel 會判定檔案損毀。
    """
    element = etree.Element(f"{{{NS_MAIN}}}drawing")
    element.set(f"{{{NS_REL}}}id", rel_id)
    _insert_before(sheet_root, element, _AFTER_DRAWING)


def _insert_legacy_drawing_hf_element(sheet_root, rel_id: str) -> None:
    """在 sheet XML 的合法位置插入 ``<legacyDrawingHF r:id=.../>``。

    須排在 ``<legacyDrawing>`` 之後、``drawingHF``／``tableParts`` 等元素
    之前（design 決策 4）。
    """
    element = etree.Element(f"{{{NS_MAIN}}}legacyDrawingHF")
    element.set(f"{{{NS_REL}}}id", rel_id)
    _insert_before(sheet_root, element, _AFTER_LEGACY_DRAWING_HF)


def _insert_before(sheet_root, element, after_local_names) -> None:
    """把 ``element`` 插入 ``sheet_root``，排在第一個屬於
    ``after_local_names`` 的子元素之前；都沒有則附加在最後。
    """
    after = {f"{{{NS_MAIN}}}{name}" for name in after_local_names}
    for index, child in enumerate(sheet_root):
        if child.tag in after:
            sheet_root.insert(index, element)
            return
    sheet_root.append(element)


def _ensure_drawing_override(parts: Dict[str, bytes], drawing_part: str) -> None:
    """補上新建 drawing part 的 Content_Types Override（沒有才新增）。"""
    name = "[Content_Types].xml"
    if name not in parts:
        return
    root = etree.fromstring(parts[name])
    overrides = {el.get("PartName") for el in root.iter(f"{{{NS_CT}}}Override")}
    if f"/{drawing_part}" in overrides:
        return
    el = etree.SubElement(root, f"{{{NS_CT}}}Override")
    el.set("PartName", f"/{drawing_part}")
    el.set("ContentType", CT_DRAWING)
    parts[name] = _serialize(root)


def _ensure_media_content_type(parts: Dict[str, bytes], ext: str) -> None:
    """補上媒體副檔名的 Content_Types Default（沒有且副檔名已知才新增）。"""
    ext = ext.lower()
    if ext not in _MEDIA_CONTENT_TYPES:
        return
    name = "[Content_Types].xml"
    if name not in parts:
        return
    root = etree.fromstring(parts[name])
    defaults = {el.get("Extension", "").lower() for el in root.iter(f"{{{NS_CT}}}Default")}
    if ext in defaults:
        return
    el = etree.SubElement(root, f"{{{NS_CT}}}Default")
    el.set("Extension", ext)
    el.set("ContentType", _MEDIA_CONTENT_TYPES[ext])
    parts[name] = _serialize(root)


def _empty_rels_root():
    return etree.Element(f"{{{NS_PKG_REL}}}Relationships",
                         nsmap={None: NS_PKG_REL})


def _ensure_vml_content_type(parts: Dict[str, bytes]) -> None:
    """補上 ``[Content_Types].xml`` 缺少的 ``vml`` Default（design 決策 4）。"""
    name = "[Content_Types].xml"
    if name not in parts:
        return
    root = etree.fromstring(parts[name])
    defaults = {el.get("Extension", "").lower() for el in root.iter(f"{{{NS_CT}}}Default")}
    if "vml" in defaults:
        return
    el = etree.SubElement(root, f"{{{NS_CT}}}Default")
    el.set("Extension", "vml")
    el.set("ContentType", CT_VML)
    parts[name] = _serialize(root)


def _restore_header_footer_vml(item: CapturedDrawing, sheet_part: str,
                               parts: Dict[str, bytes], new_parts: Dict[str, bytes]) -> None:
    """回填頁首頁尾 VML（design 決策 4）：VML 部件取空號、內容原樣寫入；
    VML 自己的 rels 重寫（``Id`` 不變、``Target`` 經 :func:`_place_media`
    取空號，VML 內的 ``o:relid`` 因此不需改寫）；工作表 rels 新增
    ``vmlDrawing`` 關聯；``<legacyDrawingHF>`` 依結構描述順序插入工作表
    XML；補 ``vml`` 內容型別。
    """
    index = 1
    while (f"xl/drawings/vmlDrawing{index}.vml" in parts
           or f"xl/drawings/vmlDrawing{index}.vml" in new_parts):
        index += 1
    vml_part = f"xl/drawings/vmlDrawing{index}.vml"
    new_parts[vml_part] = item.hf_vml

    if item.hf_vml_rels:
        vml_rels_root = _empty_rels_root()
        for rel in etree.fromstring(item.hf_vml_rels):
            added = etree.SubElement(vml_rels_root, f"{{{NS_PKG_REL}}}Relationship")
            added.set("Id", rel.get("Id"))
            added.set("Type", rel.get("Type"))
            if rel.get("Type") == REL_TYPE_IMAGE and rel.get("TargetMode") != "External":
                source_target = _resolve(vml_part, rel.get("Target", ""))
                ext = posixpath.splitext(source_target)[1].lstrip(".")
                added.set("Target", _place_media(
                    parts, new_parts, item.hf_media[source_target], ext, vml_part,
                ))
            else:
                # 外部連結：見 _relink_anchors 同一段落的說明。
                added.set("Target", rel.get("Target"))
                if rel.get("TargetMode"):
                    added.set("TargetMode", rel.get("TargetMode"))
        new_parts[_rels_path(vml_part)] = _serialize(vml_rels_root)

    sheet_rels_part = _rels_path(sheet_part)
    rels_root = (etree.fromstring(parts[sheet_rels_part])
                 if sheet_rels_part in parts else _empty_rels_root())
    rel_id = _next_free_rel_id(rels_root)
    rel = etree.SubElement(rels_root, f"{{{NS_PKG_REL}}}Relationship")
    rel.set("Id", rel_id)
    rel.set("Type", REL_TYPE_VML)
    rel.set("Target", posixpath.relpath(vml_part, posixpath.dirname(sheet_part)))
    parts[sheet_rels_part] = _serialize(rels_root)

    sheet_root = etree.fromstring(parts[sheet_part])
    _insert_legacy_drawing_hf_element(sheet_root, rel_id)
    parts[sheet_part] = _serialize(sheet_root)

    _ensure_vml_content_type(parts)


def _serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                          standalone=True)


def restore_drawings(xlsx_path: str, captured: Dict[object, CapturedDrawing],
                     shift_tracking: Optional[Dict[object, object]] = None) -> None:
    """把留存的模板圖形回填進已寫出的 xlsx，並依插列量平移錨點。

    Args:
        xlsx_path: openpyxl 剛寫出的檔案路徑（就地改寫）。
        captured: :func:`capture_unsupported_drawings` 的結果，經呼叫端
            （``Book``）轉為以載入時的 openpyxl ``Worksheet`` 物件為鍵
            （design 決策 7）：以物件當下的 ``title`` 找輸出部件，逃生門
            改名後仍能正確回填。呼叫端須先以物件身分（而非名稱字串）排除
            已從 ``workbook.worksheets`` 移除的物件才傳入本函式——僅靠
            「移除後 title 在輸出檔對不到部件」不足以保證正確略過：一旦
            另一張工作表改名重用同一個 title，字串比對會誤判為同一張
            工作表而回填到錯誤位置（見 ``Book._restore_drawings``）。
        shift_tracking: ``{Worksheet 物件: ShiftInfo}``；缺項視為不位移。

    Raises:
        Exception: 回填失敗時原樣上拋（design 決策 6：fail fast）。呼叫端
            （``Book._restore_drawings``，於 ``_atomic_save`` 的
            ``post_process`` 階段執行）負責清除暫存檔並包成 ``RenderError``，
            輸出路徑維持呼叫前狀態。
    """
    if not captured:
        return
    shift_tracking = shift_tracking or {}

    with zipfile.ZipFile(xlsx_path) as zf:
        infos = zf.infolist()
        parts = {info.filename: zf.read(info.filename) for info in infos}

    names = list(parts)
    read = parts.__getitem__
    sheet_parts = _sheet_part_map(names, read)
    new_parts: Dict[str, bytes] = {}

    for worksheet, item in captured.items():
        sheet_name = worksheet.title
        sheet_part = sheet_parts.get(sheet_name)
        if not sheet_part or sheet_part not in parts:
            logger.debug("輸出檔找不到工作表 %r，略過圖形回填", sheet_name)
            continue

        if item.anchors:
            anchors = [_shift_anchor(a, shift_tracking.get(worksheet))
                       for a in item.anchors]

            existing = _sheet_drawing_part(sheet_part, names, read)
            if existing and existing in parts:
                # openpyxl 已為此表寫出 drawing（含圖片／圖表）→ 合併進去。
                # 一張工作表只能有一個 <drawing>，不能另開一份。
                drawing_root = etree.fromstring(parts[existing])
                rels_part = _rels_path(existing)
                rels_root = (etree.fromstring(parts[rels_part])
                             if rels_part in parts else _empty_rels_root())
                anchors = _relink_anchors(anchors, item, parts, new_parts,
                                          rels_root, existing)
                for anchor_xml in anchors:
                    drawing_root.append(etree.fromstring(anchor_xml))
                parts[existing] = _serialize(drawing_root)
                if len(rels_root):
                    parts[rels_part] = _serialize(rels_root)
            else:
                # 此表沒有 drawing → 新建一份
                index = 1
                while (f"xl/drawings/drawing{index}.xml" in parts
                       or f"xl/drawings/drawing{index}.xml" in new_parts):
                    index += 1
                drawing_part = f"xl/drawings/drawing{index}.xml"

                drawing_rels_root = _empty_rels_root()
                anchors = _relink_anchors(anchors, item, parts, new_parts,
                                          drawing_rels_root, drawing_part)

                root = etree.Element(f"{{{NS_XDR}}}wsDr", nsmap={"xdr": NS_XDR})
                for anchor_xml in anchors:
                    root.append(etree.fromstring(anchor_xml))
                new_parts[drawing_part] = _serialize(root)

                if len(drawing_rels_root):
                    new_parts[_rels_path(drawing_part)] = _serialize(drawing_rels_root)

                sheet_rels_part = _rels_path(sheet_part)
                rels_root = (etree.fromstring(parts[sheet_rels_part])
                             if sheet_rels_part in parts else _empty_rels_root())
                rel_id = _next_free_rel_id(rels_root)
                rel = etree.SubElement(rels_root, f"{{{NS_PKG_REL}}}Relationship")
                rel.set("Id", rel_id)
                rel.set("Type", REL_TYPE_DRAWING)
                rel.set("Target", posixpath.relpath(drawing_part,
                                                    posixpath.dirname(sheet_part)))
                parts[sheet_rels_part] = _serialize(rels_root)

                sheet_root = etree.fromstring(parts[sheet_part])
                _insert_drawing_element(sheet_root, rel_id)
                parts[sheet_part] = _serialize(sheet_root)

                _ensure_drawing_override(parts, drawing_part)

            logger.debug("工作表 %r 回填 %d 個模板圖形", sheet_name, len(anchors))

        if item.hf_vml is not None:
            _restore_header_footer_vml(item, sheet_part, parts, new_parts)
            logger.debug("工作表 %r 回填頁首頁尾 VML", sheet_name)

    # 先依原順序寫回既有 part，再補上本次新增的 part。
    # 注意 parts 也可能長出新鍵（例如工作表原本沒有 rels，為了掛上
    # drawing 而新建的 _rels/*.rels），漏寫會讓 r:id 指向不存在的關聯，
    # openpyxl 讀得起來但 Excel 會直接拒絕開啟。
    written = set()
    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for info in infos:
            zf.writestr(info, parts[info.filename])
            written.add(info.filename)
        for name, payload in {**parts, **new_parts}.items():
            if name not in written:
                zf.writestr(name, payload)
                written.add(name)
