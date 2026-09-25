"""
容器資料模型
"""
from dataclasses import dataclass
from typing import Any, Dict, List

from .objects import Block, ObjectInfo


@dataclass
class Container:
    """
    容器資料模型
    
    包含一個Sheet中所有物件和區塊的資訊 
    """
    container_id: str
    sheet_name: str
    blocks: List[Block]
    objects: List[ObjectInfo]
    
    def to_dict(self) -> Dict[str, Any]:
        """
        轉換為字典格式
        
        Returns:
            Dict[str, Any]: 字典格式的容器資料 
        """
        return {
            "container_id": self.container_id,
            "sheet_name": self.sheet_name,
            "blocks": [block.to_dict() for block in self.blocks],
            "objects": [obj.to_dict() for obj in self.objects]
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Container':
        """
        從字典建立容器物件
        
        Args:
            data: 字典格式的容器資料
            
        Returns:
            Container: 容器物件
        """
        return cls(
            container_id=data["container_id"],
            sheet_name=data["sheet_name"],
            blocks=[Block.from_dict(block_data) for block_data in data["blocks"]],
            objects=[ObjectInfo.from_dict(obj_data) for obj_data in data["objects"]]
        )
    
    def get_block_by_id(self, block_id: str) -> Block:
        """
        根據ID取得區塊
        
        Args:
            block_id: 區塊ID
            
        Returns:
            Block: 區塊物件
            
        Raises:
            ValueError: 找不到指定的區塊
        """
        for block in self.blocks:
            if block.block_id == block_id:
                return block
        raise ValueError(f"找不到區塊ID: {block_id}")
    
    def get_object_by_id(self, obj_id: str) -> ObjectInfo:
        """
        根據ID取得物件
        
        Args:
            obj_id: 物件ID
            
        Returns:
            ObjectInfo: 物件資訊
            
        Raises:
            ValueError: 找不到指定的物件
        """
        for obj in self.objects:
            if obj.obj_id == obj_id:
                return obj
        raise ValueError(f"找不到物件ID: {obj_id}")
    
    def get_objects_by_block_id(self, block_id: str) -> List[ObjectInfo]:
        """
        根據區塊ID取得該區塊的所有物件
        
        Args:
            block_id: 區塊ID
            
        Returns:
            List[ObjectInfo]: 物件清單
        """
        return [obj for obj in self.objects if obj.block_id == block_id]
