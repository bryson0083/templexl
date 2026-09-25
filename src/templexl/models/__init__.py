"""
Excel Template Renderer - 資料模型 

包含所有核心資料模型定義：
- 基礎模型：枚舉、位置、形狀等
- 標籤模型：Simple Tag和Table Tag  
- 物件模型：物件資訊和關係模型 
- 容器模型：Container和Block模型
"""

# 基礎模型
from .base import (
    ObjectType, BlockType, DataShape, CellPosition, RangePosition
)

# 標籤模型
from .tag import Tag

# 物件模型
from .objects import ObjectInfo, ImageObject

# 容器模型
from .container import Container, Block

__all__ = [
    # 基礎模型
    'ObjectType', 'BlockType', 'DataShape', 'CellPosition', 'RangePosition',
    
    # 標籤模型
    'Tag',
    
    # 物件模型
    'ObjectInfo', 'ImageObject',
    
    # 容器模型
    'Container', 'Block',
]
