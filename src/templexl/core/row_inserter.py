"""
模板列複製與插入（row shifting）。

表格標籤渲染前，以批次方式為 N 列資料擴增空間：複製模板列的
樣式/公式/合併儲存格，單次 insert_rows 插入，再更新受影響的
標籤位置與 gap block 範圍。
"""
import logging
from typing import Any, Dict, List, Tuple

import pandas as pd
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..context import RenderContext
from ..models.base import BlockType, ObjectType
from ..models.container import Container
from .cell_ops import (
    apply_cell_style_info,
    copy_cell_style_fast,
    copy_cell_style_info,
    translate_formula,
)

logger = logging.getLogger(__name__)



def clear_original_tag_positions(worksheet: Worksheet,
    table_tags: List[Dict[str, Any]]
) -> None:
    """
    清除所有原始標籤位置和相關區域，避免模板複製時產生殘留內容

    Args:
        worksheet: Excel工作表
        table_tags: 表格標籤信息列表
    """
    logger.debug("DEBUG_CLEAR_ORIGINAL: 開始清除所有原始標籤位置和相關區域")

    cleared_positions = set()  # 避免重複清除同一位置

    for tag_info in table_tags:
        tag = tag_info['tag']
        # 使用 row_position 作為原始位置（這是標籤的初始位置）
        original_row = tag_info['row_position']
        original_col = tag.cell_position.col

        # 清除標籤位置
        position_key = (original_row, original_col)
        if position_key not in cleared_positions:
            try:
                original_cell = worksheet.cell(row=original_row, column=original_col)
                if original_cell.value:
                    logger.debug(f"清除原始標籤位置 ({original_row}, {original_col})")
                    original_cell.value = None
                    cleared_positions.add(position_key)
            except Exception as e:
                logger.debug(f"DEBUG_CLEAR_ORIGINAL: 清除標籤位置錯誤: {e}")

        # 擴展清除：清除標籤所在表格物件的所有位置，避免殘留
        # 表格物件起始位置 = 標籤位置 - 1
        table_start_row = max(1, original_row - 1)
        table_end_row = original_row

        logger.debug(f"DEBUG_CLEAR_ORIGINAL: 清除表格區域 行{table_start_row}-{table_end_row}, 列{original_col}-{original_col+5}")

        # 清除表格區域（擴展到可能的列範圍）
        for row in range(table_start_row, table_end_row + 1):
            for col in range(original_col, original_col + 6):  # 清除6列範圍
                try:
                    cell = worksheet.cell(row=row, column=col)
                    if cell.value and (str(cell.value).startswith('#{{') or
                                     str(cell.value).startswith('欄') or
                                     '欄' in str(cell.value)):
                        logger.debug(f"清除殘留內容 ({row}, {col})")
                        cell.value = None
                except Exception as e:
                    continue

    logger.debug(f"DEBUG_CLEAR_ORIGINAL: 完成清除，總共處理了 {len(cleared_positions)} 個標籤位置")


def update_positions_after_row_insertion(container: Container,
    insert_row: int,
    additional_rows: int,
    render_context: RenderContext,
    table_tags: List[Dict[str, Any]],
    target_column: int = None,
    current_tag_info: Dict[str, Any] = None
) -> None:
    """
    更新插入行之後，位於插入點下方的標籤和物件的位置
    現在考慮原始模板間距，確保表格間距正確

    Args:
        container: 容器物件
        insert_row: 插入行的位置
        additional_rows: 插入的行數
        render_context: 渲染上下文
        table_tags: 所有表格標籤的信息列表
        target_column: 可選，只影響指定列的物件（避免水平方向互相影響）
        current_tag_info: 當前標籤信息，包含原始間距
    """
    if target_column is not None:
        logger.debug(f"DEBUG: 更新插入行 {insert_row} 之後的位置，插入了 {additional_rows} 行，只影響列 {target_column}")
    else:
        logger.debug(f"DEBUG: 更新插入行 {insert_row} 之後的位置，插入了 {additional_rows} 行（所有列）")

    # 更新容器中所有物件的位置
    for obj in container.objects:
        # 如果指定了目標列，只處理該列的「表格」物件——insert_rows 實體推移
        # 的是所有欄，位於插入點下方的純量（SIMPLE）標籤必須跟著平移，
        # 不受欄過濾（否則追蹤位置與實體位置脫鉤，標籤替換會落在舊位置）
        if (
            target_column is not None
            and obj.obj_type != ObjectType.SIMPLE
            and obj.cell_position.col != target_column
        ):
            continue

        # 只更新位於插入點下方（行號較大）的物件
        if obj.cell_position.row <= insert_row:
            # 不需要更新，位置不受影響
            continue
        elif obj.cell_position.row > insert_row:
            # 需要向下移動
            old_row = obj.cell_position.row
            obj.cell_position.row += additional_rows
            logger.debug(f"DEBUG: 更新物件 {obj.obj_id} 位置: 行 {old_row} -> {obj.cell_position.row}")

            # 同時更新對應的標籤位置
            tag = render_context.get_tag_for_object(obj.obj_id)
            if tag:
                # 重要：創建標籤副本以避免共享同一個物件
                from copy import deepcopy
                tag_copy = deepcopy(tag)
                # 修正：標籤位置應該按相同偏移量移動，而不是設置為表格物件位置
                old_tag_row = tag_copy.cell_position.row
                if old_tag_row > insert_row:
                    tag_copy.cell_position.row = old_tag_row + additional_rows
                    render_context.tag_mapping[obj.obj_id] = tag_copy
                    logger.debug(f"DEBUG: 更新標籤 {tag.tag_name} 位置: 行 {old_tag_row} -> {tag_copy.cell_position.row}")
                else:
                    logger.debug(f"DEBUG: 標籤 {tag.tag_name} 位置 {old_tag_row} 不需要更新（位於插入點之前）")

    # 更新table_tags列表中的row_position（用於後續處理）
    for tag_info in table_tags:
        # 如果指定了目標列，只處理該列的標籤
        if target_column is not None and tag_info['obj'].cell_position.col != target_column:
            continue

        if tag_info['row_position'] > insert_row:
            old_pos = tag_info['row_position']
            tag_info['row_position'] += additional_rows
            logger.debug(f"DEBUG: 更新table_tags中 {tag_info['tag'].tag_name} 的記錄位置: {old_pos} -> {tag_info['row_position']}")

    # 重要：位置更新後，需要重新分配物件到正確的Block中
    logger.debug(f"DEBUG: 位置更新完成，開始重新分配物件到正確的Block")
    reassign_objects_to_correct_blocks(container)


def update_gap_block_ranges_after_insertions(container: Container, 
    tag_shape_info: Dict[str, Dict[str, Any]]
) -> None:
    """
    根據插入的行數更新Gap blocks的範圍，確保它們能包含被移動的物件

    Args:
        container: 容器物件
        tag_shape_info: 標籤shape資訊
    """
    logger.debug("DEBUG: 開始更新Gap blocks範圍以包含插入的行")

    # 計算每個Gap block中標籤插入的總行數
    for block in container.blocks:
        if block.block_type == BlockType.GAP:
            block_objects = container.get_objects_by_block_id(block.block_id)
            gap_block_expansion = 0

            for obj in block_objects:
                if obj.obj_type in [ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                    # 找到對應的標籤shape資訊
                    matching_tag_name = None
                    for tag_name, shape_info in tag_shape_info.items():
                        if shape_info['obj_info'].obj_id == obj.obj_id:
                            matching_tag_name = tag_name
                            break

                    if matching_tag_name:
                        # 優先使用實際插入的行數（已考慮header條件調整）
                        if 'actual_additional_rows' in tag_shape_info[matching_tag_name]:
                            additional_rows = tag_shape_info[matching_tag_name]['actual_additional_rows']
                            logger.debug(f"DEBUG: Gap Block {block.block_id} 中的標籤 {matching_tag_name} 使用實際插入行數: {additional_rows} 行")
                        else:
                            # 回退到原始計算方式（用於向後相容）
                            additional_rows = tag_shape_info[matching_tag_name]['rows'] - tag_shape_info[matching_tag_name]['original_rows']
                            logger.debug(f"DEBUG: Gap Block {block.block_id} 中的標籤 {matching_tag_name} 使用計算行數: {additional_rows} 行")
                        gap_block_expansion += additional_rows

            if gap_block_expansion > 0:
                original_end = block.rng_to.row
                new_end = original_end + gap_block_expansion
                logger.debug(f"DEBUG: 更新Gap Block {block.block_id} 範圍: {block.rng_from.row}-{original_end} -> {block.rng_from.row}-{new_end}")
                block.rng_to.row = new_end
            else:
                logger.debug(f"DEBUG: Gap Block {block.block_id} 不需要擴展範圍")


def copy_template_row_and_insert_new_rows(worksheet: Worksheet,
    template_row: int,
    additional_rows: int,
    tag_name: str,
) -> None:
    """
    複製模板行的樣式和公式，並插入新行

    Args:
        worksheet: Excel工作表
        template_row: 模板行號
        additional_rows: 需要插入的額外行數
        tag_name: 標籤名稱，用於跳過標籤cell本身
    """
    if additional_rows <= 0:
        return

    try:
        from openpyxl.cell.cell import MergedCell

        # logger.debug(f"DEBUG: 在第 {template_row} 行後插入 {additional_rows} 行")

        # 1. 先收集插入位置下方的合併儲存格，準備推移
        merged_ranges_to_shift = []
        for merged_range in list(worksheet.merged_cells.ranges):
            if merged_range.min_row > template_row:
                # 收集需要推移的合併儲存格
                merged_ranges_to_shift.append({
                    'original_range': str(merged_range),
                    'min_row': merged_range.min_row,
                    'max_row': merged_range.max_row,
                    'min_col': merged_range.min_col,
                    'max_col': merged_range.max_col
                })
                # 先移除原有的合併儲存格
                worksheet.unmerge_cells(str(merged_range))
                logger.debug(f"DEBUG: 暫時移除合併儲存格: {merged_range}")

        # 2. 插入所需的新行數（一次性插入，從template_row+1開始）
        worksheet.insert_rows(template_row + 1, additional_rows)

        # 3. 重新創建推移後的合併儲存格
        for merge_info in merged_ranges_to_shift:
            new_min_row = merge_info['min_row'] + additional_rows
            new_max_row = merge_info['max_row'] + additional_rows
            new_range = f"{worksheet.cell(row=new_min_row, column=merge_info['min_col']).coordinate}:{worksheet.cell(row=new_max_row, column=merge_info['max_col']).coordinate}"
            worksheet.merge_cells(new_range)
            logger.debug(f"DEBUG: 重新創建推移後的合併儲存格: {merge_info['original_range']} -> {new_range}")

        # 複製template row的樣式和公式到新插入的行
        # 效能：max_column 為模板寬度、在此迴圈內不會改變；預先計算一次，
        # 避免每列都重新掃描整張(不斷變大的)工作表 → O(n²) 降為 O(n)
        max_col = worksheet.max_column

        # 模板列每一欄「要複製什麼」在列間完全不變，於迴圈外解析一次成複製計畫，
        # 避免逐列重新取得同一批模板儲存格並重算標籤/公式判斷（熱迴圈：列數 × 欄數）。
        # 無樣式且無值可複製的欄一律排除——複製後不會產生任何效果（openpyxl 寫檔時
        # 略過無值無樣式的儲存格，見 worksheet/_writer.py write_row），故排除與否輸出等價。
        # 明細類「模板列僅一個標籤、無樣式」的情境因此整段迴圈免跑。
        copy_plan = []
        for col in range(1, max_col + 1):
            template_cell = worksheet.cell(row=template_row, column=col)
            if isinstance(template_cell, MergedCell):
                continue

            value = template_cell.value
            # 標籤儲存格只複製樣式、不複製值（標籤由渲染階段填入資料）
            is_tag_cell = (
                isinstance(value, str) and '{{' in value and '}}' in value
            )
            value_to_copy = None if is_tag_cell else value
            is_formula = (
                value_to_copy is not None
                and template_cell.data_type == 'f'
                and isinstance(value_to_copy, str)
            )
            has_style = getattr(template_cell, 'has_style', False)

            if not has_style and value_to_copy is None:
                continue

            copy_plan.append((col, template_cell, value_to_copy, is_formula))

        for copy_index in range(additional_rows):
            target_row = template_row + 1 + copy_index

            # 複製template row的內容到新插入的行
            for col, template_cell, value_to_copy, is_formula in copy_plan:
                target_cell = worksheet.cell(row=target_row, column=col)

                # 跳過合併儲存格
                if isinstance(target_cell, MergedCell):
                    continue

                # 複製cell樣式
                copy_cell_style_fast(template_cell, target_cell)

                if value_to_copy is not None:
                    if is_formula:
                        # 調整公式中的引用
                        target_cell.value = translate_formula(
                            value_to_copy,
                            template_row, col,
                            target_row, col,
                        )
                    else:
                        # 普通值（但不是標籤）
                        target_cell.value = value_to_copy

        # 複製合併儲存格
        copy_merged_cells_to_new_rows(worksheet, template_row, template_row + 1, additional_rows)

    except Exception as e:
        logger.error(f"ERROR: 複製template row失敗: {str(e)}")
        raise RenderError(f"複製template row失敗: {str(e)}")


def copy_merged_cells_to_new_rows(worksheet: Worksheet, 
    template_row: int, 
    start_target_row: int, 
    num_rows: int
) -> None:
    """
    複製模板行的合併儲存格設定到新插入的行
    只複製確實屬於數據行的合併儲存格，不複製Footer行的合併儲存格

    Args:
        worksheet: Excel工作表
        template_row: 模板行號
        start_target_row: 開始目標行號
        num_rows: 行數
    """
    try:
        # 找到確實屬於模板行本身的合併儲存格
        merged_ranges_to_copy = []
        for merged_range in worksheet.merged_cells.ranges:
            if merged_range.min_row == merged_range.max_row == template_row:
                # 只複製單行且確實屬於模板行的合併儲存格
                merged_ranges_to_copy.append(merged_range)
                logger.debug(f"DEBUG: 找到模板行 {template_row} 的合併儲存格: {merged_range}")

        # 為每個新行創建對應的合併範圍
        for i in range(num_rows):
            target_row = start_target_row + i
            for merged_range in merged_ranges_to_copy:
                # 計算新的合併範圍
                new_range = f"{worksheet.cell(row=target_row, column=merged_range.min_col).coordinate}:{worksheet.cell(row=target_row, column=merged_range.max_col).coordinate}"
                worksheet.merge_cells(new_range)
                logger.debug(f"DEBUG: 創建數據行合併儲存格: {new_range}")

                # 複製合併儲存格的樣式到所有儲存格
                # 特別注意：合併後的儲存格只有左上角儲存格可以設定值和樣式
                source_cell = worksheet.cell(row=template_row, column=merged_range.min_col)
                target_cell = worksheet.cell(row=target_row, column=merged_range.min_col)
                copy_cell_style_fast(source_cell, target_cell)

    except Exception as e:
        logger.debug(f"DEBUG: 複製合併儲存格失敗: {str(e)}")


def reassign_objects_to_correct_blocks(container: Container) -> None:
    """
    重新分配物件到正確的Block中

    在標籤位置更新後，物件可能不再屬於原來的Block範圍，
    需要重新分配到正確的Block中以確保能被正確渲染。 

    Args:
        container: 容器物件
    """
    logger.debug(f"DEBUG: 開始重新分配物件到正確的Block中")

    # 先輸出所有Block的範圍
    logger.debug(f"DEBUG: 容器的Block範圍:")
    for i, block in enumerate(container.blocks):
        logger.debug(f"DEBUG:   {i+1}. {block.block_type.value} Block {block.block_id}: 第{block.rng_from.row}-{block.rng_to.row}行")

    # 收集所有需要重新分配的物件
    reassignment_needed = []

    for obj in container.objects:
        current_row = obj.cell_position.row
        current_block_id = obj.block_id

        logger.debug(f"DEBUG: 檢查物件 {obj.obj_id} (位置: {current_row}, Block: {current_block_id})")

        # 找到該物件應該屬於的正確Block
        correct_block = None
        for block in container.blocks:
            if block.rng_from.row <= current_row <= block.rng_to.row:
                correct_block = block
                break

        if correct_block:
            logger.debug(f"DEBUG: 物件 {obj.obj_id} 應該屬於Block {correct_block.block_id}")
        else:
            logger.debug(f"DEBUG: 物件 {obj.obj_id} 沒有找到合適的Block (位置: {current_row})")
            # 找出最接近的Block
            closest_blocks = []
            for block in container.blocks:
                if current_row < block.rng_from.row:
                    closest_blocks.append(f"Block {block.block_id} 從第{block.rng_from.row}行開始 (物件在其前面)")
                elif current_row > block.rng_to.row:
                    closest_blocks.append(f"Block {block.block_id} 到第{block.rng_to.row}行結束 (物件在其後面)")
            if closest_blocks:
                logger.debug(f"DEBUG:   最接近的Block: {', '.join(closest_blocks)}")

        if correct_block and correct_block.block_id != current_block_id:
            reassignment_needed.append({
                'obj': obj,
                'old_block_id': current_block_id,
                'new_block_id': correct_block.block_id
            })
            logger.debug(f"DEBUG: 物件 {obj.obj_id} 需要重新分配：從 {current_block_id} -> {correct_block.block_id} (位置: {current_row})")

    # 執行重新分配
    for assignment in reassignment_needed:
        obj = assignment['obj']
        old_block_id = assignment['old_block_id']
        new_block_id = assignment['new_block_id']

        obj.block_id = new_block_id
        logger.debug(f"DEBUG: 完成物件 {obj.obj_id} 重新分配：{old_block_id} -> {new_block_id}")

    logger.debug(f"DEBUG: 重新分配完成，共處理 {len(reassignment_needed)} 個物件")
