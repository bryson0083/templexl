"""
物件相關資料模型
"""
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from .base import ObjectType, BlockType, CellPosition, DataShape, RangePosition


@dataclass
class ObjectInfo:
    """
    物件資訊資料模型
    
    代表Container中註冊的物件資訊 
    """
    obj_id: str
    display_name: str
    obj_type: ObjectType
    block_id: str
    sheet_name: str
    is_multi_rows: bool
    having_header: bool
    cell_position: CellPosition
    data_shape: DataShape
    obj_name: Optional[str] = None  # 原始物件名稱（用於表格物件綁定）
    
    def to_dict(self) -> Dict[str, Any]:
        """
        轉換為字典格式
        
        Returns:
            Dict[str, Any]: 字典格式的物件資料
        """
        result = {
            "obj_id": self.obj_id,
            "display_name": self.display_name,
            "obj_type": self.obj_type.value,
            "block_id": self.block_id,
            "sheet_name": self.sheet_name,
            "is_multi_rows": self.is_multi_rows,
            "having_header": self.having_header,
            "cell_position": self.cell_position.to_dict(),
            "data_shape": self.data_shape.to_dict()
        }
        # 若obj_name為None，則使用display_name
        result["obj_name"] = self.obj_name if self.obj_name is not None else self.display_name
        return result
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ObjectInfo':
        """
        從字典建立物件資訊
        
        Args:
            data: 字典格式的物件資料
            
        Returns:
            ObjectInfo: 物件資訊
        """
        # 取得obj_name，若不存在則使用display_name
        obj_name = data.get("obj_name")
        if obj_name is None:
            obj_name = data.get("display_name")
            
        return cls(
            obj_id=data["obj_id"],
            display_name=data["display_name"],
            obj_type=ObjectType(data["obj_type"]),
            block_id=data["block_id"],
            sheet_name=data["sheet_name"],
            is_multi_rows=data["is_multi_rows"],
            having_header=data["having_header"],
            cell_position=CellPosition.from_dict(data["cell_position"]),
            data_shape=DataShape.from_dict(data["data_shape"]),
            obj_name=obj_name
        )


@dataclass
class Block:
    """
    區塊資料模型
    
    代表Header、Gap、Footer區塊
    """
    block_id: str
    block_type: BlockType
    rng_from: RangePosition
    rng_to: RangePosition
    
    def to_dict(self) -> Dict[str, Any]:
        """
        轉換為字典格式
        
        Returns:
            Dict[str, Any]: 字典格式的區塊資料
        """
        return {
            "block_id": self.block_id,
            "block_type": self.block_type.value,
            "rng_from": self.rng_from.to_dict(),
            "rng_to": self.rng_to.to_dict()
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Block':
        """
        從字典建立區塊物件
        
        Args:
            data: 字典格式的區塊資料
            
        Returns:
            Block: 區塊物件
        """
        return cls(
            block_id=data["block_id"],
            block_type=BlockType(data["block_type"]),
            rng_from=RangePosition.from_dict(data["rng_from"]),
            rng_to=RangePosition.from_dict(data["rng_to"])
        )


@dataclass
class ImageObject:
    """
    圖片物件資料模型
    
    代表Excel中的圖片物件及其錨點資訊
    """
    image_id: str
    anchor_type: str  # "OneCellAnchor" 或 "TwoCellAnchor"
    from_position: RangePosition
    to_position: Optional[RangePosition] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """
        轉換為字典格式
        
        Returns:
            Dict[str, Any]: 字典格式的圖片物件資料
        """
        return {
            "image_id": self.image_id,
            "anchor_type": self.anchor_type,
            "from_position": self.from_position.to_dict(),
            "to_position": self.to_position.to_dict() if self.to_position else None
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ImageObject':
        """
        從字典建立圖片物件 
        
        Args:
            data: 字典格式的資料
            
        Returns:
            ImageObject: 圖片物件實例
        """
        return cls(
            image_id=data["image_id"],
            anchor_type=data["anchor_type"],
            from_position=RangePosition.from_dict(data["from_position"]),
            to_position=RangePosition.from_dict(data["to_position"]) 
                       if data.get("to_position") else None
        )
