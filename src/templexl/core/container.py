"""
容器管理器
"""
import logging

logger = logging.getLogger(__name__)

import uuid
from typing import Dict, Iterable, List, Optional

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from ..models.base import BlockType, ObjectType, CellPosition, DataShape, RangePosition
from ..models.container import Container
from ..models.objects import Block, ObjectInfo, ImageObject
from ..models.tag import Tag
from .template_scanner import TemplateScanner


class ContainerManager:
    """
    容器管理器類別
    
    負責建立和管理Container物件，包含物件註冊表功能 
    現在使用TemplateScanner進行統一的掃描與註冊
    """
    
    def __init__(self):
        """初始化容器管理器"""
        self.template_scanner = TemplateScanner()
    
    def create_containers(
        self,
        workbook: Workbook,
        tags: Optional[List[Tag]] = None,
        sheet_names: Optional[Iterable[str]] = None,
    ) -> List[Container]:
        """
        建立容器物件清單，每個Sheet對應一個Container

        Args:
            workbook: Excel工作簿物件
            tags: 解析出的標籤清單（可選，如果未提供將自動掃描）
            sheet_names: 限縮建立容器的工作表名稱；``None``（預設）處理全部工作表

        Returns:
            List[Container]: 容器物件清單
        """
        # 使用TemplateScanner進行統一掃描與註冊
        return self.template_scanner.scan_and_register_template(workbook, sheet_names)
    
    def classify_blocks(self, container: Container) -> Container:
        """
        將物件分類到Header/Gap/Footer區塊
        
        Args:
            container: 容器物件
            
        Returns:
            Container: 更新後的容器物件
        """
        # 找出所有表格類型的標籤位置（#前綴）
        table_positions = []
        for obj in container.objects:
            if obj.obj_type == ObjectType.TABLE:
                table_positions.append(obj.cell_position.row)
        
        table_positions.sort()
        
        # 根據表格位置分類區塊
        if not table_positions:
            # 沒有表格標籤，所有物件歸類為Header Block
            header_block = Block(
                block_id=f"header_{container.container_id}",
                block_type=BlockType.HEADER,
                rng_from=RangePosition(row=1, col=1, row_off=0, col_off=0),
                rng_to=RangePosition(row=1048576, col=16384, row_off=0, col_off=0)  # Excel最大範圍
            )
            container.blocks = [header_block]
            
            # 設定所有物件的block_id
            for obj in container.objects:
                obj.block_id = header_block.block_id
                
        else:
            # 有表格標籤，進行區塊分類
            blocks = []
            
            # Header Block: 第1行到第一個表格標籤前
            first_table_row = table_positions[0]
            if first_table_row > 1:
                header_block = Block(
                    block_id=f"header_{container.container_id}",
                    block_type=BlockType.HEADER,
                    rng_from=RangePosition(row=1, col=1, row_off=0, col_off=0),
                    rng_to=RangePosition(row=first_table_row - 1, col=16384, row_off=0, col_off=0)
                )
                blocks.append(header_block)
            
            # Gap Blocks: 表格標籤之間的區域
            for i, table_row in enumerate(table_positions):
                gap_block = Block(
                    block_id=f"gap_{i}_{container.container_id}",
                    block_type=BlockType.GAP,
                    rng_from=RangePosition(row=table_row, col=1, row_off=0, col_off=0),
                    rng_to=RangePosition(row=table_row, col=16384, row_off=0, col_off=0)  # 初始範圍
                )
                blocks.append(gap_block)
            
            # Footer Block: 最後一個表格標籤後到Sheet末尾  
            last_table_row = table_positions[-1]
            footer_block = Block(
                block_id=f"footer_{container.container_id}",
                block_type=BlockType.FOOTER,
                rng_from=RangePosition(row=last_table_row + 1, col=1, row_off=0, col_off=0),
                rng_to=RangePosition(row=1048576, col=16384, row_off=0, col_off=0)
            )
            blocks.append(footer_block)
            
            container.blocks = blocks
            
            # 分配物件到對應的區塊
            self._assign_objects_to_blocks(container)
        
        return container
    
    def _assign_objects_to_blocks(self, container: Container) -> None:
        """
        將物件分配到對應的區塊
        
        Args:
            container: 容器物件
        """
        for obj in container.objects:
            obj_row = obj.cell_position.row
            
            # 找到物件所屬的區塊
            for block in container.blocks:
                if block.rng_from.row <= obj_row <= block.rng_to.row:
                    obj.block_id = block.block_id
                    break
    
    def update_object_properties(self, container: Container, obj_id: str, properties: Dict) -> None:
        """
        更新物件屬性
        
        Args:
            container: 容器物件
            obj_id: 物件ID
            properties: 要更新的屬性字典
            
        Raises:
            ValueError: 找不到指定的物件
        """
        obj = container.get_object_by_id(obj_id)
        
        # 更新支援的屬性
        if 'data_shape' in properties:
            shape_data = properties['data_shape']
            obj.data_shape = DataShape(rows=shape_data['rows'], cols=shape_data['cols'])
        
        if 'cell_position' in properties:
            pos_data = properties['cell_position']
            obj.cell_position = CellPosition(row=pos_data['row'], col=pos_data['col'])
    
    def calculate_range_coordinates(self, container: Container) -> None:
        """
        計算所有物件的範圍座標
        
        Args:
            container: 容器物件
        """
        # 重新計算每個區塊的範圍
        for block in container.blocks:
            block_objects = container.get_objects_by_block_id(block.block_id)
            if block_objects:
                # 找出區塊中物件的最小和最大位置
                min_row = min(obj.cell_position.row for obj in block_objects)
                max_row = max(obj.cell_position.row + obj.data_shape.rows - 1 for obj in block_objects)
                min_col = min(obj.cell_position.col for obj in block_objects)
                max_col = max(obj.cell_position.col + obj.data_shape.cols - 1 for obj in block_objects)
                
                # 更新區塊範圍
                block.rng_from = RangePosition(row=min_row, col=min_col, row_off=0, col_off=0)
                block.rng_to = RangePosition(row=max_row, col=max_col, row_off=0, col_off=0)
    
    def update_container(self, container: Container, data: dict) -> None:
        """
        更新容器物件的資料形狀資訊
        
        當渲染過程中，table/table_obj 標籤完成渲染後，需要更新對應物件的 data_shape 資訊，
        以便後續的位移計算能正確進行。
        
        Args:
            container: 要更新的容器物件
            data: 包含標籤名稱和對應 DataFrame 的資料字典
        """
        import pandas as pd
        
        for obj_info in container.objects:
            # 只更新 table 和 table_obj 類型的物件
            if obj_info.obj_type in [ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                # 尋找對應的資料
                tag_name = obj_info.display_name
                
                # 移除條件後綴（例如：tb4_routeid_df | noheader -> tb4_routeid_df）
                clean_tag_name = tag_name.split('|')[0].strip() if '|' in tag_name else tag_name
                
                if clean_tag_name in data:
                    dataframe = data[clean_tag_name]
                    
                    if isinstance(dataframe, pd.DataFrame):
                        # 更新資料形狀
                        obj_info.data_shape.rows = len(dataframe)
                        obj_info.data_shape.cols = len(dataframe.columns)
                        
                        # 如果有 header，總行數需要加1
                        if obj_info.having_header:
                            obj_info.data_shape.rows += 1
                            
                        logger.debug(f"DEBUG: 更新物件 {obj_info.display_name} 的資料形狀: "
                              f"rows={obj_info.data_shape.rows}, cols={obj_info.data_shape.cols}")
