"""
區塊管理器 - 實現相同sheet下的block分區渲染搬移機制 
"""
import logging
import re

logger = logging.getLogger(__name__)

from typing import List, Dict, Any, Tuple
import pandas as pd
from copy import deepcopy

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.drawing.image import Image
from openpyxl.worksheet.table import Table

from ..models.base import BlockType, ObjectType
from ..models.container import Container
from ..models.objects import Block, ObjectInfo
from ..core.renderer import TemplateRenderer
from ..core.cell_ops import (
    apply_cell_style_info,
    copy_cell_style_info,
    write_data_value,
)
from ..core import table_writer
from .image_manager import compute_shift_info, update_image_positions_after_rendering
from ..core.row_inserter import (
    clear_original_tag_positions,
    copy_template_row_and_insert_new_rows,
    update_gap_block_ranges_after_insertions,
    update_positions_after_row_insertion,
)
from ..context import RenderContext
from ..exceptions import RenderError


class BlockManager:
    """
    區塊管理器類別
    
    負責處理Header、Gap、Footer區塊的渲染和推移，實現以下機制：
    1. 取得所有模板標籤的shape資訊，依據渲染排序計算新座標位置
    2. 以block為單位進行剪下貼上搬移（包含風格樣式、公式、圖片物件）
    3. 搬移順序從最底下開始：footer block -> gap block（註冊表排序大到小）
    4. 單一block中表格標籤渲染前，先複製template row並插入新rows
    """
    
    def __init__(self, escape_formulas: bool = True) -> None:
        """初始化BlockManager"""
        self.shape_info_cache = {}  # 快取標籤的shape資訊
        self._escape_formulas = escape_formulas
    
    def process_container_with_block_moving(
        self, 
        container: Container, 
        workbook: Workbook, 
        render_context: RenderContext,
        renderer: TemplateRenderer
    ) -> None:
        """
        使用新的block搬移機制處理容器渲染 - 流程對齊版本
        
        流程：
        1. 取得所有模板標籤的shape資訊後，依據模板標籤的渲染排序(由小到大)
        2. 針對table、table_obj的模板標籤，依據渲染排序(由小到大)，根據dataframe shape rows資訊
        3. 先複製模板標籤所在row的(包含風格樣式、公式)，再往下插入數量為(dataframe shape rows - 1)新row
        4. 並套用所複製的風格樣式、公式，藉此新增row推移下方的所有row做垂直方向的推移 
        
        Args:
            container: 容器物件
            workbook: Excel工作簿
            render_context: 渲染上下文
            renderer: 渲染器
        """
        # 確認方法被調用
        logger.debug(f"DEBUG_CONFIRM: process_container_with_block_moving called for {container.sheet_name}")
        
        worksheet = workbook[container.sheet_name]
        
        logger.debug(f"DEBUG: Starting container processing - {container.sheet_name} - aligned flow version")
        
        # 第一階段：收集所有模板標籤的shape資訊
        tag_shape_info = self._collect_all_tag_shape_info(container, render_context, workbook)
        logger.debug(f"DEBUG: 收集到 {len(tag_shape_info)} 個標籤的shape資訊")
        logger.debug(f"DEBUG_TAG_INFO: Container {container.sheet_name} collected {len(tag_shape_info)} tags")
        
        # 保存shape資訊供後續圖片位置更新使用
        self.shape_info_cache = tag_shape_info
        logger.debug(f"DEBUG_CACHE: shape_info_cache saved with {len(self.shape_info_cache)} items for {container.sheet_name}")
        
        # 第二階段：依據模板標籤的渲染排序(由小到大)，處理template row複製和插入
        self._process_template_rows_by_render_order(container, worksheet, tag_shape_info, render_context)
        logger.debug(f"DEBUG: 完成template row複製和插入")
        
        # 第三階段：執行實際的數據渲染
        self._render_all_blocks_content(container, workbook, render_context, renderer)
        logger.debug(f"DEBUG: 完成數據渲染")
    
    def _collect_all_tag_shape_info(
        self, 
        container: Container, 
        render_context: RenderContext,
        workbook: Workbook
    ) -> Dict[str, Dict[str, Any]]:
        """
        收集所有模板標籤的shape資訊
        
        Args:
            container: 容器物件
            render_context: 渲染上下文
            
        Returns:
            Dict: 標籤shape資訊 {tag_name: {rows: int, cols: int, obj_info: ObjectInfo}}
        """
        tag_shape_info = {}
        
        logger.debug(f"DEBUG_COLLECT: Container {container.sheet_name} has {len(container.objects)} objects")
        
        for obj in container.objects:
            logger.debug(f"DEBUG_OBJ: Object {obj.obj_id} type: {obj.obj_type.value if hasattr(obj.obj_type, 'value') else obj.obj_type}")
            if obj.obj_type in [ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                logger.debug(f"DEBUG_TABLE: Found table object {obj.obj_id}")
                tag = render_context.get_tag_for_object(obj.obj_id)
                logger.debug(f"DEBUG_TAG: Tag for {obj.obj_id}: {tag}")
                if tag:
                    has_data = render_context.has_data(tag.tag_name)
                    logger.debug(f"DEBUG_HAS_DATA: tag.tag_name={tag.tag_name}, has_data={has_data}")
                    if has_data:
                        data = render_context.get_data(tag.tag_name)
                        logger.debug(f"DEBUG_DATA: Got data type={type(data)}")
                        
                        if hasattr(data, 'shape'):  # DataFrame
                            data_rows = data.shape[0]
                            data_cols = data.shape[1]
                            
                            # 是否包含header行
                            total_rows = data_rows
                            
                            # 對於 TABLE_OBJ 類型，檢查表格物件本身是否有標題行
                            if obj.obj_type == ObjectType.TABLE_OBJ:
                                # TABLE_OBJ 的表格物件本身可能有標題行，需要保留空間
                                # 即使有 noheader 條件也要計算標題行空間
                                worksheet = workbook[obj.sheet_name]
                                if worksheet.tables and obj.obj_name:
                                    for table in worksheet.tables.values():
                                        if table.displayName == obj.obj_name:
                                            if table.headerRowCount > 0:
                                                total_rows += 1
                                            break
                            elif obj.having_header and not (tag.has_condition and tag.condition == "noheader"):
                                # 對於其他類型（TABLE），維持原邏輯
                                total_rows += 1
                            
                            tag_shape_info[tag.tag_name] = {
                                'rows': total_rows,
                                'cols': data_cols,
                                'original_rows': obj.data_shape.rows,  # 使用實際的原始行數
                                'obj_info': obj,
                                'tag': tag
                            }
                            
                            logger.debug(f"DEBUG: 標籤 {tag.tag_name} shape: {total_rows}行 x {data_cols}列")
                            logger.debug(f"DEBUG_COLLECTED: Added {tag.tag_name} to shape_info")
                        else:
                            logger.debug(f"DEBUG: 標籤 {tag.tag_name} 的數據不是DataFrame: {type(data)}")
                else:
                    if tag:
                        logger.debug(f"DEBUG: 標籤 {tag.tag_name} 沒有對應的數據")
                    else:
                        logger.debug(f"DEBUG: 物件 {obj.obj_id} 沒有對應的標籤")
            elif obj.obj_type == ObjectType.SIMPLE:
                # 處理簡單標籤（不需要shape資訊，但要確保能被渲染）
                tag = render_context.get_tag_for_object(obj.obj_id)
                if tag:
                    logger.debug(f"DEBUG: 發現簡單標籤 {tag.tag_name} 在位置 ({tag.cell_position.row}, {tag.cell_position.col})")
                    if render_context.has_data(tag.tag_name):
                        logger.debug(f"DEBUG: 簡單標籤 {tag.tag_name} 有對應數據")
                    else:
                        logger.debug(f"DEBUG: 簡單標籤 {tag.tag_name} 沒有對應數據")
        
        return tag_shape_info
        # DEBUG: Image position update completed

    # 保持現有方法的兼容性

    def _process_template_rows_by_render_order(
        self, 
        container: Container, 
        worksheet: Worksheet,
        tag_shape_info: Dict[str, Dict[str, Any]],
        render_context: RenderContext
    ) -> None:
        """
        依據模板標籤的渲染排序(由小到大)，處理template row複製和插入
        
        對每個table、table_obj的模板標籤：
        1. 根據dataframe shape rows資訊
        2. 先複製模板標籤所在row的(包含風格樣式、公式)
        3. 再往下插入數量為(dataframe shape rows - 1)新row
        4. 並套用所複製的風格樣式、公式
        5. 藉此新增row推移下方的所有row做垂直方向的推移
        
        Args:
            container: 容器物件
            worksheet: Excel工作表
            tag_shape_info: 標籤shape資訊
            render_context: 渲染上下文
        """
        logger.debug("DEBUG: 開始依據渲染排序處理template row複製和插入")
        
        # 收集所有table、table_obj的標籤，並按照渲染排序(由小到大，即row位置由上到下)
        table_tags = []
        
        for obj in container.objects:
            if obj.obj_type in [ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                tag = render_context.get_tag_for_object(obj.obj_id)
                if tag and tag.tag_name in tag_shape_info:
                    table_tags.append({
                        'tag': tag,
                        'obj': obj,
                        'shape_info': tag_shape_info[tag.tag_name],
                        'row_position': tag.cell_position.row
                    })
        
        # 修正：按列分組然後排序，避免水平方向不同列的表格互相影響位置
        # 重要：只考慮垂直方向推移，水平方向分開處理
        def group_by_columns_then_sort(table_tags):
            """
            按列位置分組，然後在每組內按行位置排序（由大到小，從下到上）
            這樣可以確保水平方向不同列的表格不會互相影響位置計算
            同時計算每組內表格間的原始間距
            """
            # 按列位置分組
            column_groups = {}
            for tag_info in table_tags:
                col = tag_info['obj'].cell_position.col
                if col not in column_groups:
                    column_groups[col] = []
                column_groups[col].append(tag_info)

            # 在每組內按行位置排序並計算間距
            result = []
            for col in sorted(column_groups.keys()):
                group = column_groups[col]
                # 先按行位置排序（從小到大，從上到下）
                group.sort(key=lambda x: x['row_position'])

                # 計算每個表格與前一個表格的原始間距 - 修正版
                for i, tag_info in enumerate(group):
                    if i == 0:
                        # 第一個表格，無需計算間距
                        tag_info['original_gap'] = 0
                    else:
                        # 修正：正確計算表格物件間的gap
                        prev_tag_info = group[i-1]

                        # 前一個表格的結束位置 = 標籤位置 (因為標籤在表格最後一行)
                        prev_end_row = prev_tag_info['row_position']

                        # 當前表格的起始位置 = 標籤位置 - 1 (因為標籤在表格第2行)
                        curr_table_start_row = tag_info['row_position'] - 1

                        # 原始間距 = 當前表格起始位置 - 前表格結束位置 - 1
                        original_gap = curr_table_start_row - prev_end_row - 1
                        tag_info['original_gap'] = max(0, original_gap)

                        logger.debug(f"DEBUG: 標籤 {tag_info['tag'].tag_name} gap計算: 前表格結束{prev_end_row} -> 當前表格開始{curr_table_start_row} = gap {original_gap}行")

                # 反轉順序（由大到小，從下到上）用於處理
                group.reverse()
                result.extend(group)
                logger.debug(f"DEBUG: 列 {col} 包含 {len(group)} 個標籤")
                for tag_info in group:
                    logger.debug(f"DEBUG:   列 {col} 標籤 {tag_info['tag'].tag_name} 在第 {tag_info['row_position']} 行")

            return result

        # 使用新的分組排序邏輯
        table_tags = group_by_columns_then_sort(table_tags)

        logger.debug(f"DEBUG: 找到 {len(table_tags)} 個需要處理的table標籤")
        for tag_info in table_tags:
            logger.debug(f"DEBUG:   標籤 {tag_info['tag'].tag_name} 在第 {tag_info['row_position']} 行")

        # 重要：將計算的gap信息轉移到objects中，供後續渲染時使用
        logger.debug("DEBUG: 轉移gap信息到objects中")
        for tag_info in table_tags:
            obj = tag_info['obj']
            original_gap = tag_info.get('original_gap', 0)
            # 將gap信息儲存到object的自定義屬性中
            setattr(obj, '_temp_gap', original_gap)
            logger.debug(f"DEBUG: 標籤 {tag_info['tag'].tag_name} 的gap信息已轉移: {original_gap} 行")

        # 重要：在開始處理前，先清除所有原始標籤位置，避免模板複製時殘留標籤
        logger.debug("DEBUG: 預先清除所有原始標籤位置，避免模板複製時產生殘留內容")
        clear_original_tag_positions(worksheet, table_tags)
        
        # 修正：維護按列分組的累積位移量字典，避免水平方向表格互相影響
        # 格式：{column: {row: shift_amount}}
        column_cumulative_shifts = {}
        
        logger.debug("DEBUG: 開始從下往上處理template row複製和插入（避免位置互相影響）")
        
        for tag_info in table_tags:
            tag = tag_info['tag']
            obj = tag_info['obj']
            shape_info = tag_info['shape_info']
            # 使用標籤物件的當前位置，而不是初始記錄的位置
            template_row = tag.cell_position.row
            
            # 計算需要插入的行數
            total_rows_needed = shape_info['rows']
            original_rows = shape_info['original_rows']

            # 按照實際需求計算插入行數（僅數據行）
            additional_rows = total_rows_needed - original_rows

            # 修正策略：不在插入階段處理gap，而是在渲染時調整位置
            original_gap = tag_info.get('original_gap', 0)

            # 重要：將實際插入的行數記錄到shape_info中，供Gap Block擴展使用
            shape_info['actual_additional_rows'] = additional_rows  # 修正：只記錄數據行插入

            # 將gap信息保存到shape_info中，供後續渲染時使用
            shape_info['original_gap'] = original_gap
            
            logger.debug(f"DEBUG: 處理標籤 {tag.tag_name}:")
            logger.debug(f"DEBUG:   當前template_row: {template_row}")
            logger.debug(f"DEBUG:   初始記錄位置: {tag_info['row_position']}")
            logger.debug(f"DEBUG:   total_rows_needed: {total_rows_needed}")
            logger.debug(f"DEBUG:   original_rows: {original_rows}")
            logger.debug(f"DEBUG:   additional_rows: {additional_rows} (僅數據行)")
            logger.debug(f"DEBUG:   original_gap: {original_gap} (將在渲染時處理)")

            if additional_rows > 0:
                logger.debug(f"DEBUG: 標籤 {tag.tag_name} 需要插入 {additional_rows} 行數據行，gap在渲染時處理")

                # 複製模板行的樣式和公式，並插入新行（僅數據行）
                copy_template_row_and_insert_new_rows(
                    worksheet,
                    template_row,
                    additional_rows,  # 只插入數據行
                    tag.tag_name,  # 傳入標籤名稱，用於跳過標籤本身
                )

                # 記錄此次插入的累積位移（按列分組）- 只記錄數據行插入
                col = obj.cell_position.col
                if col not in column_cumulative_shifts:
                    column_cumulative_shifts[col] = {}
                column_cumulative_shifts[col][template_row] = additional_rows
                logger.debug(f"DEBUG: 記錄列 {col} 行 {template_row} 的累積位移: {additional_rows}")

                # 修正：只更新同一列中位於此標籤下方的標籤和物件的位置
                # 避免水平方向不同列的表格互相影響位置
                # gap間距將在後續渲染時處理
                update_positions_after_row_insertion(
                    container,
                    template_row,
                    additional_rows,  # 修正：只使用數據行插入量
                    render_context,
                    table_tags,  # 傳入所有表格標籤信息
                    col,  # 新增：指定只影響此列
                    tag_info  # 傳入當前標籤信息，包含原始間距
                )
                
                # 更新物件的data_shape資訊
                obj.data_shape.rows = total_rows_needed
                obj.data_shape.cols = shape_info['cols']
                
                logger.debug(f"DEBUG: 標籤 {tag.tag_name} 處理完成，插入了 {additional_rows} 行數據行，gap:{original_gap}行將在渲染時處理")
            else:
                logger.debug(f"DEBUG: 標籤 {tag.tag_name} 不需要插入額外行數，gap:{original_gap}行將在渲染時處理")
        
        logger.debug("DEBUG: template row複製和插入處理完成")
        
        # 重要：更新所有Gap blocks的範圍以包含插入的行
        update_gap_block_ranges_after_insertions(container, tag_shape_info)
    
    def _render_all_blocks_content(
        self,
        container: Container,
        workbook: Workbook,
        render_context: RenderContext,
        renderer: TemplateRenderer
    ) -> None:
        """
        渲染所有區塊的內容（包括簡單標籤和表格標籤）
        
        在行插入完成後，執行實際的數據渲染
        
        Args:
            container: 容器物件
            workbook: Excel工作簿
            render_context: 渲染上下文
            renderer: 渲染器
        """
        worksheet = workbook[container.sheet_name]
        
        logger.debug(f"DEBUG: 開始渲染所有區塊內容 - {container.sheet_name}")
        
        # 收集所有需要渲染的標籤
        tags_to_render = []
        
        for obj in container.objects:
            tag = render_context.get_tag_for_object(obj.obj_id)
            if tag and render_context.has_data(tag.tag_name):
                tags_to_render.append({
                    'tag': tag,
                    'obj': obj,
                    'data': render_context.get_data(tag.tag_name)
                })
                logger.debug(f"DEBUG: 準備渲染標籤 {tag.tag_name} 在位置 ({tag.cell_position.row}, {tag.cell_position.col})")
        
        logger.debug(f"DEBUG: 總共 {len(tags_to_render)} 個標籤需要渲染")
        
        # 按位置排序（從上到下，從左到右）
        tags_to_render.sort(key=lambda x: (x['tag'].cell_position.row, x['tag'].cell_position.col))
        
        # 渲染每個標籤
        for tag_info in tags_to_render:
            tag = tag_info['tag']
            obj = tag_info['obj']
            data = tag_info['data']
            
            logger.debug(f"渲染標籤 {tag.tag_name} - 類型: {obj.obj_type.value} - 位置: ({tag.cell_position.row}, {tag.cell_position.col}) - 資料型別: {type(data).__name__}")
            
            if obj.obj_type == ObjectType.SIMPLE:
                # 渲染簡單標籤
                self._render_simple_tag(worksheet, tag, data, renderer)
            elif obj.obj_type in [ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                # 渲染表格標籤
                table_writer.render_table_tag(
                    worksheet, tag, obj, data, renderer, container, render_context,
                    escape_formulas=self._escape_formulas,
                    shape_info_cache=self.shape_info_cache,
                )
            
        logger.debug(f"DEBUG: 完成所有標籤渲染")
        
        # 渲染完成後，更新圖片物件位置
        # DEBUG: Preparing to update image positions
        update_image_positions_after_rendering(container, workbook, self.shape_info_cache)
        # DEBUG: Image position update completed

        # 記錄本表的插列位移，供存檔後回填模板圖形（core.drawing_keeper）平移錨點。
        # openpyxl 不模型化 <xdr:sp>，圖形無法留在 workbook 物件裡跟著位移，
        # 只能在 zip 層回填時依這份位移量重算錨點。
        shift_info = compute_shift_info(container, self.shape_info_cache)
        if shift_info is not None:
            render_context.shift_tracking[container.sheet_name] = shift_info
    
    def _render_simple_tag(
        self,
        worksheet: Worksheet,
        tag,
        data: Any,
        renderer=None
    ) -> None:
        """
        渲染簡單標籤 - 委託給renderer進行適當的處理
        
        Args:
            worksheet: Excel工作表
            tag: 標籤物件
            data: 要渲染的數據
            renderer: 渲染器實例（可選）
        """
        # 如果有renderer，使用renderer的專業方法處理（支持DataFrame等複雜類型）
        if renderer:
            try:
                # 創建一個臨時workbook參考（renderer需要）
                workbook = worksheet.parent
                renderer.render_simple_tag(tag, data, workbook, worksheet)
                # logger.debug(f"DEBUG: 使用renderer渲染簡單標籤 {tag.tag_name} 在 ({tag.cell_position.row}, {tag.cell_position.col})")
                return
            except Exception as e:
                logger.debug(f"DEBUG: renderer渲染失敗，回退到基本方法: {e}")
        
        # 回退到基本的字符串替換方法
        row = tag.cell_position.row
        col = tag.cell_position.col
        
        # 獲取儲存格
        cell = worksheet.cell(row=row, column=col)
        
        # 檢查是否為合併儲存格
        if isinstance(cell, MergedCell):
            logger.debug(f"DEBUG: 跳過合併儲存格 ({row}, {col})")
            return
        
        # 獲取當前儲存格的值
        current_value = cell.value
        
        if current_value and isinstance(current_value, str):
            # 使用正規表達式替換標籤為實際數據，支援標籤內的空白字符
            # 構建支援任意空白的標籤模式：{{ tag_name }} 或 {{tag_name}}
            tag_pattern = rf"\{{\{{\s*{re.escape(tag.tag_name)}\s*\}}\}}"
            
            # 檢查是否包含標籤
            if re.search(tag_pattern, current_value):
                # 使用正規表達式替換標籤為數據值
                new_value = re.sub(tag_pattern, str(data), current_value)
                if current_value.startswith('='):
                    # 模板作者的公式（含標籤），保留公式語意
                    cell.value = new_value
                else:
                    write_data_value(cell, new_value, self._escape_formulas)
                logger.debug(f"簡單標籤 {tag.tag_name} 已渲染於 ({row}, {col})")
            else:
                logger.debug(f"警告 - 儲存格 ({row}, {col}) 不包含標籤模式 {tag_pattern}")
        else:
            # 如果儲存格沒有值或不是字符串，直接設置數據
            write_data_value(cell, data, self._escape_formulas, str_fallback=False)
            logger.debug(f"簡單標籤 {tag.tag_name} 直接設置值於 ({row}, {col})")
