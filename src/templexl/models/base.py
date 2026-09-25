"""
基礎資料類型和常量定義 
"""
from enum import Enum
from typing import Any, Dict, List, Optional
from dataclasses import dataclass


class TagType(Enum):
    """標籤類型"""
    SIMPLE = "simple"
    TABLE = "table"


class DataType(Enum):
    """數據類型"""
    STRING = "string"
    NUMBER = "number"
    DATE = "date"
    DATAFRAME = "dataframe"
    UNKNOWN = "unknown"


class RenderDirection(Enum):
    """渲染方向"""
    HORIZONTAL = "horizontal"
    VERTICAL = "vertical"


class BlockType(Enum):
    """區塊類型"""
    HEADER = "Header"
    GAP = "Gap"
    FOOTER = "Footer"


class ObjectType(Enum):
    """物件類型"""
    SIMPLE = "simple"
    TABLE = "table"
    TABLE_OBJ = "table_obj"
    IMAGE_OBJ = "image_obj"
    TEXT = "text"


@dataclass
class CellPosition:
    """儲存格位置"""
    row: int
    col: int
    
    def to_dict(self) -> Dict[str, Any]:
        """轉換為字典格式"""
        return {"row": self.row, "col": self.col}
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CellPosition':
        """從字典建立物件"""
        return cls(row=data["row"], col=data["col"])


@dataclass
class RangePosition:
    """範圍位置"""
    row: int
    col: int
    row_off: int
    col_off: int
    
    def to_dict(self) -> Dict[str, Any]:
        """轉換為字典格式"""
        return {
            "row": self.row,
            "col": self.col,
            "rowOff": self.row_off,
            "colOff": self.col_off
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'RangePosition':
        """從字典建立物件"""
        return cls(
            row=data["row"],
            col=data["col"],
            row_off=data.get("rowOff", 0),
            col_off=data.get("colOff", 0)
        )


@dataclass
class DataShape:
    """數據形狀"""
    rows: int
    cols: int
    
    def to_dict(self) -> Dict[str, Any]:
        """轉換為字典格式"""
        return {"rows": self.rows, "cols": self.cols}
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DataShape':
        """從字典建立物件"""
        return cls(rows=data["rows"], cols=data["cols"])
