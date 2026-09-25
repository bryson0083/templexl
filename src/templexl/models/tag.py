"""
標籤相關資料模型
"""
from dataclasses import dataclass
from typing import Any, Dict, Optional

from .base import TagType, DataType, RenderDirection, CellPosition


@dataclass
class Tag:
    """
    模板標籤資料模型 
    
    代表Excel模板中的一個標籤，包含標籤名稱、類型、位置等資訊
    """
    tag_name: str
    tag_type: TagType
    has_condition: bool
    condition: Optional[str]
    cell_position: CellPosition
    data_type: DataType
    render_direction: RenderDirection
    sheet_name: str  # 新增sheet_name欄位
    
    def to_dict(self) -> Dict[str, Any]:
        """
        轉換為字典格式
        
        Returns:
            Dict[str, Any]: 字典格式的標籤資料
        """
        return {
            "tag_name": self.tag_name,
            "tag_type": self.tag_type.value,
            "has_condition": self.has_condition,
            "condition": self.condition,
            "cell_position": self.cell_position.to_dict(),
            "data_type": self.data_type.value,
            "render_direction": self.render_direction.value,
            "sheet_name": self.sheet_name
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Tag':
        """
        從字典建立標籤物件
        
        Args:
            data: 字典格式的標籤資料 
            
        Returns:
            Tag: 標籤物件
        """
        return cls(
            tag_name=data["tag_name"],
            tag_type=TagType(data["tag_type"]),
            has_condition=data["has_condition"],
            condition=data.get("condition"),
            cell_position=CellPosition.from_dict(data["cell_position"]),
            data_type=DataType(data["data_type"]),
            render_direction=RenderDirection(data["render_direction"]),
            sheet_name=data["sheet_name"]
        )
    
    def __str__(self) -> str:
        """字串表示"""
        condition_str = f" | {self.condition}" if self.has_condition else ""
        prefix = "#" if self.tag_type == TagType.TABLE else ""
        return f"{prefix}{{{{{self.tag_name}{condition_str}}}}}"
