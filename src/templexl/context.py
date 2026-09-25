"""
渲染上下文
"""
from typing import Any, Dict


class RenderContext:
    """
    渲染上下文類別
    
    包含渲染過程中需要的所有上下文資訊 
    """
    
    def __init__(self, process_id: str, template_path: str, output_path: str, data: Dict[str, Any]):
        self.process_id = process_id
        self.template_path = template_path
        self.output_path = output_path
        self.data = data
        self.shift_tracking = {}  # 用於追蹤位移資訊
        self.tag_mapping = {}  # 用於標籤映射
    
    def get_data(self, key: str) -> Any:
        """
        取得渲染資料
        
        Args:
            key: 資料鍵值
            
        Returns:
            Any: 資料值
            
        Raises:
            KeyError: 找不到指定的資料
        """
        if key not in self.data:
            raise KeyError(f"找不到渲染資料: {key}")
        return self.data[key]
    
    def has_data(self, key: str) -> bool:
        """
        檢查是否有指定的資料
        
        Args:
            key: 資料鍵值
            
        Returns:
            bool: 是否存在
        """
        return key in self.data
    
    def add_tag_mapping(self, obj_id: str, tag) -> None:
        """
        添加物件到標籤的映射
        
        Args:
            obj_id: 物件ID
            tag: 標籤物件
        """
        self.tag_mapping[obj_id] = tag
    
    def get_tag_for_object(self, obj_id: str):
        """
        取得物件對應的標籤
        
        Args:
            obj_id: 物件ID
            
        Returns:
            Tag: 標籤物件，如果找不到則返回None
        """
        return self.tag_mapping.get(obj_id)
