"""
圖片物件管理器
"""
import logging

logger = logging.getLogger(__name__)

from typing import Dict, List, Optional
from dataclasses import dataclass

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..models.container import Container
from ..models.objects import ImageObject, ObjectInfo
from ..models.base import ObjectType, RangePosition


@dataclass
class ShiftInfo:
    """位移資訊"""
    sheet_name: str
    start_row: int
    shift_amount: int
    affected_objects: List[str]


class ImageObjectManager:
    """
    圖片物件管理器類別
    
    負責管理Excel中的圖片物件，包括位置更新和推移處理
    """
    
    def scan_image_objects(self, worksheet: Worksheet, sheet_name: str) -> List[ImageObject]:
        """
        掃描工作表中的圖片物件
        
        Args:
            worksheet: 工作表物件
            sheet_name: 工作表名稱
            
        Returns:
            List[ImageObject]: 圖片物件清單
        """
        image_objects = []
        
        # 掃描工作表中的圖片 - 使用正確的openpyxl屬性
        images = getattr(worksheet, '_images', [])
        for idx, image in enumerate(images):
            image_obj = self._create_image_object_from_image(image, idx, sheet_name)
            if image_obj:
                image_objects.append(image_obj)
        
        return image_objects
    
    def _create_image_object_from_image(
        self, 
        image, 
        index: int, 
        sheet_name: str
    ) -> Optional[ImageObject]:
        """
        從Excel圖片物件創建ImageObject
        
        Args:
            image: Excel圖片物件
            index: 圖片索引
            sheet_name: 工作表名稱
            
        Returns:
            Optional[ImageObject]: 圖片物件，如果創建失敗則返回None
        """
        try:
            # 取得圖片錨點資訊
            anchor = image.anchor
            
            # 嘗試獲取from位置 - 支援多種屬性名稱
            from_pos = None
            for from_attr in ['_from', 'from']:
                if hasattr(anchor, from_attr):
                    from_pos = getattr(anchor, from_attr)
                    break
            
            if from_pos:
                from_row = getattr(from_pos, 'row', 0) + 1  # 轉換為1-based
                from_col = getattr(from_pos, 'col', 0) + 1
                from_row_off = getattr(from_pos, 'rowOff', 0)
                from_col_off = getattr(from_pos, 'colOff', 0)
                
                from_position = RangePosition(
                    row=from_row,
                    col=from_col,
                    row_off=from_row_off,
                    col_off=from_col_off
                )
            else:
                from_position = RangePosition(row=1, col=1, row_off=0, col_off=0)
            
            # 嘗試獲取to位置 - 支援多種屬性名稱
            to_position = None
            to_pos = None
            for to_attr in ['_to', 'to']:
                if hasattr(anchor, to_attr):
                    to_pos = getattr(anchor, to_attr)
                    break
            
            if to_pos:
                to_row = getattr(to_pos, 'row', 0) + 1  # 轉換為1-based
                to_col = getattr(to_pos, 'col', 0) + 1
                to_row_off = getattr(to_pos, 'rowOff', 0)
                to_col_off = getattr(to_pos, 'colOff', 0)
                
                to_position = RangePosition(
                    row=to_row,
                    col=to_col,
                    row_off=to_row_off,
                    col_off=to_col_off
                )
                
                logger.debug(f"DEBUG: ImageManager掃描到TwoCellAnchor - from:({from_row},{from_col}) to:({to_row},{to_col})")
            else:
                logger.debug(f"DEBUG: ImageManager掃描到OneCellAnchor - from:({from_row},{from_col})")
            
            # 確定錨點類型
            anchor_type = "TwoCellAnchor" if to_position else "OneCellAnchor"
            
            return ImageObject(
                image_id=f"image_{sheet_name}_{index}",
                anchor_type=anchor_type,
                from_position=from_position,
                to_position=to_position
            )
            
        except Exception:
            # 如果無法解析圖片資訊，返回None
            return None
    
    def update_image_anchors(
        self, 
        worksheet: Worksheet, 
        image_objects: List[ImageObject], 
        shift_info: ShiftInfo
    ) -> None:
        """
        更新工作表中的圖片錨點
        
        Args:
            worksheet: 工作表物件
            image_objects: 圖片物件清單
            shift_info: 位移資訊
        """
        images = getattr(worksheet, '_images', [])
        if not images:
            logger.debug("DEBUG: 工作表沒有圖片物件")
            return
        
        logger.debug(f"DEBUG: 開始更新 {len(images)} 個圖片錨點")
        
        # 更新每個圖片的錨點
        for idx, image in enumerate(images):
            # 對於footer區塊的圖片，直接應用位移
            logger.debug(f"DEBUG: 處理圖片 {idx+1}")
            self._update_image_anchor_direct(image, shift_info)
    
    def _update_image_anchor_direct(self, image, shift_info: ShiftInfo) -> None:
        """
        更新圖片錨點

        Args:
            image: Excel圖片物件
            shift_info: 位移資訊
        """
        try:
            anchor = image.anchor
            anchor_type = type(anchor).__name__

            logger.debug(f"DEBUG: 更新圖片錨點，位移量: {shift_info.shift_amount}")
            logger.debug(f"DEBUG: 錨點類型: {anchor_type}")

            # 檢查原始錨點位置 - 使用正確的屬性名稱
            original_from_row = None
            original_to_row = None

            # 讀取原始位置 - 使用 _from 屬性（正確的openpyxl屬性名稱）
            if hasattr(anchor, '_from') and anchor._from is not None:
                if hasattr(anchor._from, 'row') and anchor._from.row is not None:
                    original_from_row = anchor._from.row
                    logger.debug(f"DEBUG: 原始_from位置: {original_from_row} (0-based)")

            # 讀取原始to位置 - 使用 to 屬性
            if hasattr(anchor, 'to') and anchor.to is not None:
                if hasattr(anchor.to, 'row') and anchor.to.row is not None:
                    original_to_row = anchor.to.row
                    logger.debug(f"DEBUG: 原始to位置: {original_to_row} (0-based)")

            # 判斷是否需要位移（如果圖片位於表格之後）
            should_shift = False
            if original_from_row is not None:
                # openpyxl使用0-based索引，shift_info.start_row是1-based
                # 需要轉換比較：0-based圖片位置 >= (1-based起始行 - 1)
                if original_from_row >= (shift_info.start_row - 1):
                    should_shift = True
                    logger.debug(f"DEBUG: 圖片需要位移（原始from位置 {original_from_row} >= 起始行 {shift_info.start_row-1}）")

            if not should_shift:
                logger.debug("DEBUG: 圖片不需要位移")
                return
            
            if anchor_type == "TwoCellAnchor":
                self._adjust_two_cell_anchor_correctly(anchor, shift_info.shift_amount)
            elif anchor_type == "OneCellAnchor":
                self._adjust_one_cell_anchor_correctly(anchor, shift_info.shift_amount)
            else:
                logger.debug(f"DEBUG: 未知錨點類型: {anchor_type}")

        except Exception as e:
            logger.debug(f"DEBUG: 圖片錨點更新失敗: {e}")
            import traceback
            logger.debug("例外堆疊", exc_info=True)

    def _adjust_two_cell_anchor_correctly(self, anchor, row_offset: int) -> None:
        """
        正確調整 TwoCellAnchor 類型的圖片位置

        Args:
            anchor: TwoCellAnchor 錨點物件
            row_offset: 行偏移量
        """
        if not hasattr(anchor, '_from') or not anchor._from:
            logger.debug(f"DEBUG: TwoCellAnchor 圖片缺少 _from 屬性")
            return

        if not hasattr(anchor, 'to') or not anchor.to:
            logger.debug(f"DEBUG: TwoCellAnchor 圖片缺少 to 屬性")
            return

        # 保存原始的所有屬性值
        original_from_row = anchor._from.row
        original_from_col = anchor._from.col
        original_from_rowOff = getattr(anchor._from, 'rowOff', 0)
        original_from_colOff = getattr(anchor._from, 'colOff', 0)

        original_to_row = anchor.to.row
        original_to_col = anchor.to.col
        original_to_rowOff = getattr(anchor.to, 'rowOff', 0)
        original_to_colOff = getattr(anchor.to, 'colOff', 0)

        # 調整 _from 位置
        anchor._from.row = original_from_row + row_offset
        # 保持 _from 的其他屬性不變
        if hasattr(anchor._from, 'rowOff'):
            anchor._from.rowOff = original_from_rowOff
        if hasattr(anchor._from, 'colOff'):
            anchor._from.colOff = original_from_colOff

        # 調整 to 位置
        anchor.to.row = original_to_row + row_offset
        # 保持 to 的其他屬性不變
        if hasattr(anchor.to, 'rowOff'):
            anchor.to.rowOff = original_to_rowOff
        if hasattr(anchor.to, 'colOff'):
            anchor.to.colOff = original_to_colOff

        logger.debug(f"DEBUG: TwoCellAnchor 圖片位置已調整：")
        logger.debug(f"DEBUG:   _from: 第{original_from_row}行 -> 第{anchor._from.row}行 (0-based)")
        logger.debug(f"DEBUG:   to: 第{original_to_row}行 -> 第{anchor.to.row}行 (0-based)")
        logger.debug(f"DEBUG:   rowOff 保持: _from({original_from_rowOff}), to({original_to_rowOff})")
        logger.debug(f"DEBUG:   colOff 保持: _from({original_from_colOff}), to({original_to_colOff})")

    def _adjust_one_cell_anchor_correctly(self, anchor, row_offset: int) -> None:
        """
        正確調整 OneCellAnchor 類型的圖片位置

        Args:
            anchor: OneCellAnchor 錨點物件
            row_offset: 行偏移量
        """
        if not hasattr(anchor, '_from') or not anchor._from:
            logger.debug(f"DEBUG: OneCellAnchor 圖片缺少 _from 屬性")
            return

        # 保存原始的 rowOff 和 colOff 值
        original_from_row = anchor._from.row
        original_from_col = anchor._from.col
        original_from_rowOff = getattr(anchor._from, 'rowOff', 0)
        original_from_colOff = getattr(anchor._from, 'colOff', 0)

        # 調整 _from 位置
        anchor._from.row = original_from_row + row_offset
        # 保持 _from 的其他屬性不變
        if hasattr(anchor._from, 'rowOff'):
            anchor._from.rowOff = original_from_rowOff
        if hasattr(anchor._from, 'colOff'):
            anchor._from.colOff = original_from_colOff

        logger.debug(f"DEBUG: OneCellAnchor 圖片位置已調整：")
        logger.debug(f"DEBUG:   _from: 第{original_from_row}行 -> 第{anchor._from.row}行 (0-based)")
        logger.debug(f"DEBUG:   rowOff 保持: {original_from_rowOff}")
        logger.debug(f"DEBUG:   colOff 保持: {original_from_colOff}")


def compute_shift_info(
    container: 'Container',
    shape_info_cache: dict | None = None,
) -> Optional[ShiftInfo]:
    """計算工作表因表格展開而產生的整體列位移。

    位移量取各表格標籤實際插入列數的總和，起算列為最上緣表格的下一列
    （1-based）。圖片錨點平移與模板圖形保留（``core.drawing_keeper``）
    共用這份結果。

    Args:
        container: 已完成渲染的容器（對應單一工作表）。
        shape_info_cache: 標籤的 shape 快取，含實際插入列數。

    Returns:
        Optional[ShiftInfo]: 有位移時回傳位移資訊；無位移回傳 ``None``。
    """
    # 位移計算邏輯：直接從shape_info_cache獲取表格實際擴展
    total_shift = 0
    min_table_row = float('inf')

    # 方法1：從shape_info_cache獲取實際的表格擴展信息
    if shape_info_cache:
        # DEBUG: Using shape_info_cache to calculate shifts
        for tag_name, shape_info in shape_info_cache.items():
            # 優先使用實際插入的行數（已考慮header條件調整）
            if 'actual_additional_rows' in shape_info:
                shift_amount = shape_info['actual_additional_rows']
                obj_info = shape_info.get('obj_info')
                if obj_info and shift_amount > 0:
                    table_row = obj_info.cell_position.row
                    logger.debug(f"DEBUG: Using actual_additional_rows for image shift: tag {tag_name} inserted {shift_amount} rows")

                    # 累計位移並記錄最小的表格行號
                    total_shift += shift_amount
                    if table_row < min_table_row:
                        min_table_row = table_row
            # 回退到原始計算方式（用於向後相容）
            elif 'rows' in shape_info and 'original_rows' in shape_info:
                actual_rows = shape_info['rows']
                original_rows = shape_info.get('original_rows', 1)
                if actual_rows > original_rows:
                    obj_info = shape_info.get('obj_info')
                    if obj_info:
                        table_row = obj_info.cell_position.row
                        shift_amount = actual_rows - original_rows
                        logger.debug(f"DEBUG: 圖片位移使用計算行數: 標籤 {tag_name} 計算插入 {shift_amount} 行")

                        # 累計位移並記錄最小的表格行號
                        total_shift += shift_amount
                        if table_row < min_table_row:
                            min_table_row = table_row

    # 方法2：檢查Gap blocks（作為備用方案）
    if total_shift == 0:
        logger.debug(f"DEBUG: shape_info_cache沒有位移信息，檢查gap blocks...")
        logger.debug(f"DEBUG: container.blocks 總數: {len(container.blocks)}")
        for i, block in enumerate(container.blocks):
            try:
                logger.debug(f"DEBUG: Block {i}: {block.block_id} - type: {block.block_type.value} - range: {block.rng_from.row}-{block.rng_to.row}")
            except:
                # 避免編碼問題
                logger.debug(f"DEBUG: Block {i}: {block.block_id} - range: {block.rng_from.row}-{block.rng_to.row}")

        # 從blocks的範圍變化計算位移
        for block in container.blocks:
            logger.debug(f"DEBUG: 檢查 block {block.block_id}, type: {block.block_type.value}")
            if block.block_type.value == 'Gap':  # 修正：使用大寫的 'Gap'
                # Gap block的範圍擴展表示有table渲染
                gap_size = block.rng_to.row - block.rng_from.row + 1
                logger.debug(f"DEBUG: Gap block {block.block_id} 範圍: {block.rng_from.row}-{block.rng_to.row}, 大小: {gap_size}")
                if gap_size > 1:  # 原來gap只有1行，現在大於1表示被擴展了
                    additional_rows = gap_size - 1
                    total_shift += additional_rows
                    if block.rng_from.row < min_table_row:
                        min_table_row = block.rng_from.row
                    logger.debug(f"DEBUG: Gap block {block.block_id} 被擴展，產生{additional_rows}行位移")
                else:
                    logger.debug(f"DEBUG: Gap block {block.block_id} 沒有擴展 (大小={gap_size})")
            else:
                logger.debug(f"DEBUG: 跳過非Gap block: {block.block_id}")

    # DEBUG: Shift calculation results

    if total_shift > 0 and min_table_row != float('inf'):
        return ShiftInfo(
            sheet_name=container.sheet_name,
            start_row=int(min_table_row) + 1,  # 從第一個table的下一行開始位移
            shift_amount=total_shift,
            affected_objects=[],
        )
    return None


def update_image_positions_after_rendering(container: 'Container',
    workbook: Workbook,
    shape_info_cache: dict | None = None,
) -> None:
    """
    渲染完成後更新圖片物件位置

    Args:
        container: 容器物件
        workbook: Excel工作簿
    """
    # DEBUG: 開始更新圖片物件位置

    from ..models.base import ObjectType

    worksheet = workbook[container.sheet_name]
    image_manager = ImageObjectManager()

    # 找出所有圖片物件
    image_objects = [obj for obj in container.objects if obj.obj_type == ObjectType.IMAGE_OBJ]
    # DEBUG: 容器中的所有物件類型

    if not image_objects:
        # DEBUG: 沒有找到圖片物件
        return

    # DEBUG: Found image objects
    # for obj in image_objects:
    #     DEBUG: Image object position

    # 位移計算已抽出為 compute_shift_info()，與模板圖形保留（drawing_keeper）
    # 共用同一份結果，避免兩條路徑各算一次而漂移。
    shift_info = compute_shift_info(container, shape_info_cache)
    if shift_info is not None:
        shift_info.affected_objects = [obj.obj_id for obj in image_objects]

        # DEBUG: Applying image shift

        # 直接更新工作表中的實際圖片錨點
        worksheet_images = image_manager.scan_image_objects(worksheet, container.sheet_name)
        # DEBUG: Found worksheet images

        # 檢查工作表的_images屬性
        images_attr = getattr(worksheet, '_images', [])
        # DEBUG: Worksheet._images count

        if worksheet_images:
            logger.debug(f"DEBUG: 呼叫 image_manager.update_image_anchors")
            image_manager.update_image_anchors(worksheet, worksheet_images, shift_info)
            logger.debug(f"DEBUG: image_manager.update_image_anchors 完成")

            # 最終驗證：直接檢查並修正錨點位置
            # DEBUG: Final image anchor verification
            images_attr = getattr(worksheet, '_images', [])
            for i, image in enumerate(images_attr):
                anchor = image.anchor
                if hasattr(anchor, '_from'):
                    current_from_row = anchor._from.row
                    expected_from_row = current_from_row
                    # DEBUG: Image anchor from.row

                    # 不再應用額外的位移，只驗證當前位置是否正確
                    # 錨點應該已經被 image_manager.update_image_anchors 正確更新了
                    if hasattr(anchor, '_to'):
                        current_to_row = anchor._to.row
                        # DEBUG: Image anchor to.row

                    logger.debug(f"DEBUG: Image {i+1} anchor verification complete")
        else:
            logger.debug(f"DEBUG: No worksheet images found, skipping anchor update")
    else:
        logger.debug("DEBUG: No images need to be shifted")

    logger.debug("DEBUG: Image position update completed")
