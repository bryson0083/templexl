"""
templexl 例外階層

所有對外例外皆繼承自 ``TemplateError``，呼叫端可用它一次攔截本套件的全部錯誤。
"""


class TemplateError(Exception):
    """所有 templexl 例外的基底類別。"""
    pass


#: 向後相容別名（舊名稱），新程式請使用 ``TemplateError``。
ExcelTemplateRendererError = TemplateError


class TemplateNotFoundError(TemplateError):
    """模板檔案不存在例外"""
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        super().__init__(f"模板檔案不存在: {template_path}")


class InvalidDataTypeError(TemplateError):
    """不支援的資料類型例外"""
    
    def __init__(self, data_type: str, tag_name: str):
        self.data_type = data_type
        self.tag_name = tag_name
        super().__init__(f"標籤 '{tag_name}' 不支援資料類型: {data_type}")


class InvalidTagSyntaxError(TemplateError):
    """無效標籤語法例外"""
    
    def __init__(self, tag_string: str, position: str = ""):
        self.tag_string = tag_string
        self.position = position
        pos_info = f" (位置: {position})" if position else ""
        super().__init__(f"無效的標籤語法: '{tag_string}'{pos_info}")


class RenderError(TemplateError):
    """渲染過程錯誤例外"""
    
    def __init__(self, message: str, tag_name: str = "", sheet_name: str = ""):
        self.tag_name = tag_name
        self.sheet_name = sheet_name
        location_info = ""
        if sheet_name:
            location_info += f" (Sheet: {sheet_name}"
            if tag_name:
                location_info += f", 標籤: {tag_name}"
            location_info += ")"
        super().__init__(f"渲染錯誤: {message}{location_info}")


class RangeOverlapError(TemplateError):
    """範圍重疊錯誤例外"""
    
    def __init__(self, obj1_id: str, obj2_id: str, sheet_name: str = ""):
        self.obj1_id = obj1_id
        self.obj2_id = obj2_id
        self.sheet_name = sheet_name
        sheet_info = f" (Sheet: {sheet_name})" if sheet_name else ""
        super().__init__(f"物件範圍重疊: {obj1_id} 與 {obj2_id}{sheet_info}")


class TableObjectConflictError(TemplateError):
    """表格物件衝突錯誤例外"""
    
    def __init__(self, table_ids: list, sheet_name: str = ""):
        self.table_ids = table_ids
        self.sheet_name = sheet_name
        table_list = ", ".join(table_ids)
        sheet_info = f" (Sheet: {sheet_name})" if sheet_name else ""
        super().__init__(f"表格物件範圍衝突: {table_list}{sheet_info}")


class FileFormatError(TemplateError):
    """檔案格式錯誤例外"""

    def __init__(self, file_path: str, expected_format: str = "xlsx/xlsm"):
        self.file_path = file_path
        self.expected_format = expected_format
        super().__init__(f"不支援的檔案格式: {file_path}，期望格式: {expected_format}")


class TemplateResourceError(TemplateError):
    """模板檔資源占用超限例外（解壓炸彈防護）。"""

    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"模板檔資源檢查未通過: {file_path}（{reason}）")


class StructuralEditError(TemplateError):
    """結構操作（如刪整欄）無法保證正確性而拒絕執行。

    工作表存在本套件無法正確調整的特徵時拋出——例如公式（openpyxl 刪欄
    不改寫參照，且複製語意的平移不等於刪除語意的參照改寫）、圖片、
    Excel 表格、條件格式、資料驗證、模板形狀（AutoShape、文字方塊等
    openpyxl 無法表示的繪圖）、openpyxl 會丟棄的圖片（如 EMF／WMF）。
    與其產出靜默損壞的報表，不如明確拒絕。
    """

    def __init__(self, operation: str, reason: str, sheet_name: str = ""):
        self.operation = operation
        self.reason = reason
        self.sheet_name = sheet_name
        sheet_info = f"（工作表 '{sheet_name}'）" if sheet_name else ""
        super().__init__(f"無法執行 {operation}{sheet_info}: {reason}")


class MemoryError(TemplateError):
    """記憶體不足錯誤例外"""

    def __init__(self, operation: str = ""):
        self.operation = operation
        op_info = f" ({operation})" if operation else ""
        super().__init__(f"記憶體不足 {op_info}")
