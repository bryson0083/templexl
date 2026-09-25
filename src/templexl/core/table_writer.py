"""
DataFrame → 表格渲染與 Excel Table 物件範圍同步（block_manager 側）。

處理 #{{表格}} 標籤的資料寫入：優先綁定 Excel Table 物件並同步
table.ref/autoFilter.ref；無 Table 物件時直接渲染內容。

兩條路徑的樣式語意一致——資料列樣式唯一來源為模板列樣式複製
（row_inserter），本模組不添加模板中不存在的樣式。

表格幾何方面本模組只負責「表頭＋body」的範圍計算（body 至少一列），
合計列的延伸與篩選範圍規則一律交給 ``table_sync``；位置備援比對亦以
``expected_autofilter_ref()`` 算出的表頭＋body 範圍為準，使落在合計列的
標籤在掃描層與渲染層都不綁定表格。
"""
import logging
from typing import Any, Dict, Optional

import pandas as pd
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import coordinate_to_tuple, get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from .cell_ops import apply_cell_style_info, write_data_value
from .renderer import TemplateRenderer
from .table_sync import expected_autofilter_ref, update_table_range_sync

logger = logging.getLogger(__name__)


def render_table_tag(worksheet: Worksheet,
    tag,
    obj,
    data,
    renderer: TemplateRenderer,
    container=None,
    render_context=None,
    *,
    escape_formulas: bool = True,
    shape_info_cache: dict | None = None,
) -> None:
    """
    渲染表格標籤（DataFrame數據）

    Args:
        worksheet: Excel工作表
        tag: 標籤物件
        obj: 物件資訊
        data: DataFrame數據
        renderer: 渲染器
    """
    import pandas as pd

    if not isinstance(data, pd.DataFrame):
        logger.debug(f"DEBUG: 警告 - 標籤 {tag.tag_name} 的數據不是DataFrame")
        return

    row = tag.cell_position.row
    col = tag.cell_position.col

    # 清除標籤本身 - 需要檢查原始位置和當前位置
    original_row = tag.cell_position.row
    original_col = tag.cell_position.col

    logger.debug(f"DEBUG_TAG_CLEAR: 檢查標籤清除 - 原始位置: ({original_row}, {original_col})")

    # 清除原始位置的標籤
    original_cell = worksheet.cell(row=original_row, column=original_col)


    # 保存原始樣式
    original_style = copy_cell_style_info(original_cell) if original_cell.value else None

    # 檢查是否為合併儲存格，如果是則跳過設置值
    from openpyxl.cell.cell import MergedCell
    if not isinstance(original_cell, MergedCell):
        original_cell.value = None
        logger.debug(f"DEBUG_TAG_CLEAR: 已清除原始位置 ({original_row}, {original_col}) 的標籤")

    # 如果當前使用的位置與原始位置不同，也要清除當前位置
    if row != original_row or col != original_col:
        current_cell = worksheet.cell(row=row, column=col)
        logger.debug(f"檢查殘留標籤位置 ({row}, {col})")
        if not isinstance(current_cell, MergedCell):
            current_cell.value = None
            logger.debug(f"DEBUG_TAG_CLEAR: 已清除當前位置 ({row}, {col}) 的標籤")

    # 檢查是否需要渲染header
    render_header = obj.having_header if hasattr(obj, 'having_header') else True
    logger.debug(f"DEBUG_TAG_CONDITION: 標籤 {tag.tag_name} - has_condition: {tag.has_condition}, condition: '{tag.condition}'")
    if tag.has_condition and tag.condition == "noheader":
        render_header = False
        logger.debug(f"DEBUG_TAG_CONDITION: 標籤 {tag.tag_name} 被識別為 noheader 模式")

    logger.debug(f"DEBUG: 渲染表格標籤 {tag.tag_name} - 渲染header: {render_header}")

    # 不管物件類型是 TABLE 還是 TABLE_OBJ，都要檢查是否有表格物件需要更新
    # 找到對應的表格物件
    tables = worksheet.tables
    table_obj = None

    # 修正：優先透過物件顯示名稱匹配表格物件（更可靠）
    obj_display_name = getattr(obj, 'display_name', None)
    logger.debug(f"DEBUG: 查找表格物件 - 標籤: {tag.tag_name}, 物件顯示名稱: {obj_display_name}")

    # 首先嘗試透過顯示名稱匹配（更智能的匹配策略）
    if obj_display_name:
        # 策略1：完全匹配
        for table_name in tables:
            table = tables[table_name]
            if obj_display_name in table_name or table_name.endswith(obj_display_name):
                table_obj = table
                logger.debug(f"DEBUG: 透過名稱匹配找到表格: {table_name}")
                break

        # 策略2：如果找不到，嘗試通過標籤名稱匹配（去掉前綴）
        if not table_obj and tag.tag_name:
            for table_name in tables:
                table = tables[table_name]
                # 檢查表格名稱是否包含標籤名稱
                if tag.tag_name in table_name:
                    table_obj = table
                    logger.debug(f"DEBUG: 透過標籤名稱匹配找到表格: {table_name}")
                    break

        # 策略3：基於順序的智能匹配（針對多表格場景）
        if not table_obj:
            # 獲取所有表格，按名稱排序
            sorted_tables = sorted(tables.keys())
            logger.debug(f"DEBUG: 可用表格列表: {sorted_tables}")

            # 根據標籤名稱推斷是第幾個表格
            table_index = -1
            if 'report_df' == tag.tag_name:
                table_index = 0
            elif 'report2_df' == tag.tag_name:
                table_index = 1
            elif 'report3_df' == tag.tag_name:
                table_index = 2
            elif 'report4_df' == tag.tag_name:
                table_index = 3

            if table_index >= 0 and table_index < len(sorted_tables):
                table_name = sorted_tables[table_index]
                table_obj = tables[table_name]
                logger.debug(f"DEBUG: 透過順序匹配找到表格: {table_name} (索引: {table_index})")
            else:
                logger.debug(f"DEBUG: 順序匹配失敗 - 標籤: {tag.tag_name}, 索引: {table_index}, 可用表格數: {len(sorted_tables)}")

    # 如果名稱匹配失敗，再嘗試透過位置找到對應的表格物件（備用方案）
    if not table_obj:
        from openpyxl.utils import cell as cell_utils
        logger.debug(f"DEBUG: 名稱匹配失敗，嘗試位置匹配 - 標籤位置: ({row}, {col})")
        for table_name in tables:
            table = tables[table_name]
            logger.debug(f"DEBUG: 檢查表格 {table_name}, 範圍: {getattr(table, 'ref', 'N/A')}")
            # 解析表格範圍（只取「表頭＋資料列」——落在合計列的標籤不綁定表格，
            # 與掃描層 _bind_tags_to_tables 的判定保持一致）
            if hasattr(table, 'ref'):
                ref_str = expected_autofilter_ref(table)
                # 將A2:F2這樣的範圍字符串解析為座標
                if ':' in ref_str:
                    start_cell_str, end_cell_str = ref_str.split(':')
                    start_row, start_col = cell_utils.coordinate_to_tuple(start_cell_str)
                    end_row, end_col = cell_utils.coordinate_to_tuple(end_cell_str)

                    # 檢查標籤位置是否在表格範圍內或表格範圍附近（考慮到標籤可能在表格下方）
                    # 由於插入行後，標籤位置可能已經改變，所以要考慮原始位置關係
                    # 更精確的條件：標籤位置在表格起始行，且在表格列範圍內
                    if (row == start_row and start_col <= col <= end_col):
                        table_obj = table
                        logger.debug(f"DEBUG: 透過位置匹配找到表格: {table_name}")
                        break
                    # 或者標籤在表格範圍內的其他位置
                    elif (start_row <= row <= end_row and start_col <= col <= end_col):
                        table_obj = table
                        logger.debug(f"DEBUG: 透過位置匹配找到表格: {table_name}")
                        break

    if table_obj:
        # 渲染數據到表格物件並更新範圍
        render_dataframe_to_table_with_range_update(
            worksheet, data, table_obj, row, col, render_header, container, tag, render_context,
            escape_formulas=escape_formulas, shape_info_cache=shape_info_cache,
        )
    else:
        # 如果找不到表格物件，使用普通渲染
        logger.debug(f"DEBUG: 找不到對應的表格物件，標籤位置 ({row},{col})，使用普通渲染")
        render_dataframe_content(
            worksheet, data, row, col, render_header, original_style,
            escape_formulas=escape_formulas,
        )

    logger.debug(f"DEBUG: 完成表格標籤 {tag.tag_name} 渲染")


def _iter_normalized_rows(dataframe):
    """逐列產出已正規化缺失值的資料值 tuple。

    取代 ``DataFrame.iterrows()``：後者每列都建立一個 Series，在大表渲染的
    熱迴圈中成本可觀；改為逐欄一次性物化，再以 ``zip`` 橫向組成列。

    缺失值（``NaN`` / ``pd.NA`` / ``NaT``）於此**單點**統一轉為 ``None``，
    使下游寫值不必逐格呼叫 ``pd.isna()``。這也是缺失值語意的唯一收斂點——
    openpyxl 不接受 ``pd.NA``（拋 ValueError），而 ``str(pd.NA)`` 會在報表上
    留下 ``<NA>`` 字樣，兩者皆為實務踩過的坑。

    附帶效果：逐欄物化使各欄保有自身型別，不再受 ``iterrows()`` 逐列統一
    dtype 的影響（int 欄與 float 欄同列時，整列會被提升為 float）。
    """
    if dataframe.empty or len(dataframe.columns) == 0:
        return iter(())

    columns = []
    for _, series in dataframe.items():
        values = series.astype(object).to_numpy(dtype=object, copy=True)
        values[series.isna().to_numpy()] = None
        columns.append(values.tolist())

    return zip(*columns)


def render_dataframe_content(worksheet: Worksheet,
    dataframe,
    start_row: int,
    start_col: int,
    render_header: bool,
    original_style=None,
    *,
    escape_formulas: bool = True,
) -> None:
    """
    渲染DataFrame內容到工作表

    資料列的樣式（含框線）於本函式呼叫前，已由 row_inserter 自模板列複製；
    本函式只寫值，不添加模板中不存在的樣式。

    Args:
        worksheet: Excel工作表
        dataframe: DataFrame數據
        start_row: 起始行
        start_col: 起始列
        render_header: 是否渲染表頭
        original_style: 原始樣式
    """
    import pandas as pd

    current_row = start_row

    # 渲染表頭（如果需要）
    if render_header:
        from openpyxl.cell.cell import MergedCell
        for col_idx, column_name in enumerate(dataframe.columns):
            header_cell = worksheet.cell(row=current_row, column=start_col + col_idx)
            # 檢查是否為合併儲存格
            if not isinstance(header_cell, MergedCell):
                # 保留原始樣式
                if original_style and col_idx == 0:
                    apply_cell_style_info(header_cell, original_style)
                header_cell.value = column_name
        current_row += 1

    # 渲染數據
    from openpyxl.cell.cell import MergedCell

    # 工作表無任何合併儲存格時（明細類 raw data dump 的常態），逐格的
    # 合併範圍掃描與 MergedCell 檢查恆為徒勞——此處走緊湊路徑略過兩者，
    # 但保留「模板已預置公式則不覆蓋」的既有契約。
    # 判定依據為 merged_cells.ranges 是否為空：本函式不建立合併範圍，
    # 故該集合在整個迴圈期間不變。
    if not worksheet.merged_cells.ranges:
        for row_data in _iter_normalized_rows(dataframe):
            for col_idx, value in enumerate(row_data):
                data_cell = worksheet.cell(row=current_row, column=start_col + col_idx)
                existing_value = data_cell.value
                if isinstance(existing_value, str) and existing_value.startswith('='):
                    # 儲存格已包含公式，跳過不覆蓋，讓公式自動計算
                    continue
                write_data_value(
                    data_cell, value, escape_formulas, str_fallback=False
                )
            current_row += 1

        logger.debug(
            f"DEBUG: DataFrame內容已渲染（無合併快路徑）- 起始位置: ({start_row}, {start_col}), header: {render_header}"
        )
        return

    for row_data in _iter_normalized_rows(dataframe):
        cumulative_offset = 0  # 追蹤合併儲存格造成的累積偏移

        for col_idx, value in enumerate(row_data):
            target_col = start_col + col_idx + cumulative_offset

            # 檢查目標位置是否有合併儲存格
            merge_range = None
            for mr in worksheet.merged_cells.ranges:
                if mr.min_row <= current_row <= mr.max_row and mr.min_col <= target_col <= mr.max_col:
                    merge_range = mr
                    break

            # 如果目標位置在合併範圍內
            if merge_range:
                # 計算是否為合併範圍的左上角
                is_top_left = (current_row == merge_range.min_row and target_col == merge_range.min_col)

                if is_top_left:
                    # 這是合併範圍的左上角，正常寫入數據
                    data_cell = worksheet.cell(row=current_row, column=target_col)
                    if not isinstance(data_cell, MergedCell):
                        # 檢查儲存格是否已經包含公式，如果有則保留公式不覆蓋
                        existing_value = data_cell.value
                        if isinstance(existing_value, str) and existing_value.startswith('='):
                            # 儲存格已包含公式，跳過不覆蓋，讓公式自動計算
                            logger.debug(f"保留公式 - 位置 ({current_row}, {target_col})")
                        else:
                            write_data_value(
                                data_cell, value, escape_formulas, str_fallback=False
                            )

                    # 計算此合併儲存格佔用的額外列數
                    merge_span = merge_range.max_col - merge_range.min_col
                    cumulative_offset += merge_span
                    logger.debug(f"DEBUG: 合併儲存格偵測 - 行{current_row}, 列{target_col}, 合併範圍: {merge_range}, 新偏移: {cumulative_offset}")
                else:
                    # 不是左上角，需要跳過並調整偏移
                    # 找到合併範圍的右邊界，將目標列移到合併範圍之後
                    target_col = merge_range.max_col + 1
                    cumulative_offset = target_col - (start_col + col_idx)

                    # 在合併範圍之後寫入數據
                    data_cell = worksheet.cell(row=current_row, column=target_col)
                    if not isinstance(data_cell, MergedCell):
                        # 檢查儲存格是否已經包含公式，如果有則保留公式不覆蓋
                        existing_value = data_cell.value
                        if isinstance(existing_value, str) and existing_value.startswith('='):
                            # 儲存格已包含公式，跳過不覆蓋，讓公式自動計算
                            logger.debug(f"保留公式 - 位置 ({current_row}, {target_col})")
                        else:
                            write_data_value(
                                data_cell, value, escape_formulas, str_fallback=False
                            )
                    logger.debug(f"DEBUG: 跳過合併儲存格 - 行{current_row}, 原列{start_col + col_idx}, 新列{target_col}, 偏移: {cumulative_offset}")
            else:
                # 沒有合併儲存格，正常寫入數據
                data_cell = worksheet.cell(row=current_row, column=target_col)
                if not isinstance(data_cell, MergedCell):
                    # 檢查儲存格是否已經包含公式，如果有則保留公式不覆蓋
                    existing_value = data_cell.value
                    if isinstance(existing_value, str) and existing_value.startswith('='):
                        # 儲存格已包含公式，跳過不覆蓋，讓公式自動計算
                        logger.debug(f"保留公式 - 位置 ({current_row}, {target_col})")
                    else:
                        write_data_value(
                            data_cell, value, escape_formulas, str_fallback=False
                        )

        current_row += 1

    logger.debug(f"DEBUG: DataFrame內容已渲染 - 起始位置: ({start_row}, {start_col}), header: {render_header}")


def render_dataframe_to_table_with_range_update(worksheet: Worksheet,
    dataframe,
    table_obj,
    start_row: int,
    start_col: int,
    render_header: bool,
    container=None,
    tag=None,
    render_context=None,
    *,
    escape_formulas: bool = True,
    shape_info_cache: dict | None = None,
) -> None:
    """
    渲染DataFrame數據到表格物件並更新表格範圍

    本函式算出的範圍語意為「**表頭列＋資料列**」，交由
    ``table_sync.update_table_range_sync`` 負責延伸合計列與設定篩選範圍。
    資料列數取 ``max(len(df), 1)``——Excel 表格至少需要一列資料列，空
    DataFrame 時模板列清除標籤後物理上就是一列空白資料列。

    Args:
        worksheet: Excel工作表
        dataframe: DataFrame數據
        table_obj: Excel表格物件
        start_row: 起始行
        start_col: 起始列
        render_header: 是否渲染表頭
    """
    import pandas as pd
    from openpyxl.utils import range_boundaries
    from openpyxl.cell.cell import MergedCell

    # 獲取原始表格範圍資訊
    orig_min_col, orig_min_row, orig_max_col, orig_max_row = range_boundaries(table_obj.ref)
    original_table_cols = orig_max_col - orig_min_col + 1
    dataframe_cols = len(dataframe.columns)


    # 使用DataFrame欄位數和原始表格欄位數中的較大值
    actual_cols = max(dataframe_cols, original_table_cols)

    # 檢查原始表格是否包含表頭
    original_has_header = table_obj.headerRowCount and table_obj.headerRowCount > 0

    # 渲染數據（修正：使用表格物件的調整後位置）
    # 獲取對應的表格物件位置
    table_obj_row = orig_min_row  # 表格物件的起始行

    # 檢查表格物件位置是否被template row插入調整過
    # 查找對應的物件來獲取其當前位置
    matching_obj = None
    # 從當前容器中找到對應的表格物件
    if container and tag:
        for obj in container.objects:
            if hasattr(obj, 'display_name') and obj.display_name == tag.tag_name:
                matching_obj = obj
                break

    if matching_obj:
        # 修正：需要獲取對應標籤的實際位置，而不是表格物件位置
        # 查找對應的標籤
        corresponding_tag = render_context.get_tag_for_object(matching_obj.obj_id)
        if corresponding_tag:
            tag_row = corresponding_tag.cell_position.row
        else:
            # 回退：如果找不到標籤，使用表格物件位置+1（表格物件在標籤上方一行）
            tag_row = matching_obj.cell_position.row + 1

        # 重要：從container的shape_info_cache獲取gap信息來調整表格位置
        original_gap = 0
        if tag and hasattr(tag, 'tag_name') and shape_info_cache is not None:
            # 修正：使用正確的cache key格式（只用tag_name，不包含sheet_name）
            cache_key = tag.tag_name
            if cache_key in shape_info_cache:
                shape_info = shape_info_cache[cache_key]
                original_gap = shape_info.get('original_gap', 0)
                logger.debug(f"DEBUG: 從shape_info_cache獲取到gap信息: {original_gap} 行")
            else:
                logger.debug(f"DEBUG: 在shape_info_cache中找不到 {cache_key}")
                # DEBUG: 列印所有可用的cache keys
                logger.debug(f"DEBUG: 可用的cache keys: {list(shape_info_cache.keys())}")
        else:
            logger.debug(f"DEBUG: 無法從shape_info_cache獲取gap信息")

        # 表格起始位置計算：
        # 重要修正：template row插入過程已經維持了正確的相對位置和gap
        # 因此直接使用標籤位置-1作為表格位置，不需要額外應用gap
        actual_table_row = max(1, tag_row - 1)  # 標籤位置 - 1 = 表格起始位置
        actual_table_col = matching_obj.cell_position.col

        logger.debug(f"DEBUG: 表格位置計算: 標籤行{tag_row} -> 表格行{actual_table_row}, 列{actual_table_col} (template插入已維持gap)")
    else:
        actual_table_row = table_obj_row
        actual_table_col = orig_min_col
        logger.debug(f"DEBUG: 使用表格物件原始位置: 行{actual_table_row}, 列{actual_table_col}")

    if render_header:
        # 如果需要渲染表頭，表頭從表格物件位置開始
        current_row = actual_table_row
        current_col = actual_table_col
        logger.debug(f"DEBUG: 表頭渲染位置: ({current_row}, {current_col})")
    else:
        # noheader模式：數據從表頭下一行開始渲染，保留模板原有表頭
        # 重要修正：noheader模式下不檢查保護偏移，因為我們已經在_copy_template_row_and_insert_new_rows中跳過了保護
        current_row = actual_table_row + 1  # 只跳過表頭行，不需要額外保護偏移
        current_col = actual_table_col
        logger.debug(f"DEBUG: 數據渲染位置（noheader模式）: ({current_row}, {current_col}) - 只跳過表頭行，無保護偏移")

    # 渲染表頭（如果需要）
    if render_header:
        for col_idx, column_name in enumerate(dataframe.columns):
            if col_idx < actual_cols:
                cell = worksheet.cell(row=current_row, column=current_col + col_idx)
                if not isinstance(cell, MergedCell):
                    cell.value = column_name
        current_row += 1

    # 渲染數據行
    for row_idx, row_data in dataframe.iterrows():
        for col_idx, value in enumerate(row_data):
            if col_idx < actual_cols:
                cell = worksheet.cell(row=current_row, column=current_col + col_idx)
                if not isinstance(cell, MergedCell):
                    write_data_value(
                        cell, value, escape_formulas, str_fallback=False
                    )
        current_row += 1

    # 更新表格範圍 - 考慮原始表格是否包含表頭
    # Excel 表格至少需要一列資料列：空 DataFrame 時模板列已被清除標籤、
    # 物理上就是一列空白資料列，此處把它算進範圍（只有表頭的 ref 在 Excel 中不合法）。
    body_rows = max(len(dataframe), 1)

    # 確定表格範圍的起始位置
    if render_header:
        # 我們渲染了表頭，表格範圍從表格物件實際位置開始
        table_start_row = actual_table_row  # 表格物件位置
        table_start_col = actual_table_col
        total_rows = body_rows + 1  # 包含表頭行
        logger.debug(f"DEBUG: 有表頭模式 - 表格起始: ({table_start_row}, {table_start_col}) (表格物件位置), 總行數: {total_rows}")
    else:
        # noheader模式：表格範圍仍從表格物件位置開始，但包含保留的模板表頭
        table_start_row = actual_table_row  # 表格物件位置（包含模板表頭）
        table_start_col = actual_table_col
        total_rows = body_rows + 1  # 包含保留的模板表頭行
        logger.debug(f"DEBUG: noheader模式 - 表格起始: ({table_start_row}, {table_start_col}) (包含模板表頭), 總行數: {total_rows}")

    new_end_row = table_start_row + total_rows - 1
    new_end_col = table_start_col + actual_cols - 1

    # 格式化新的範圍字符串
    from openpyxl.utils import get_column_letter
    start_cell = f"{get_column_letter(table_start_col)}{table_start_row}"
    end_cell = f"{get_column_letter(new_end_col)}{new_end_row}"
    new_range = f"{start_cell}:{end_cell}"

    logger.debug(f"DEBUG: 表格範圍更新 - 原始: {table_obj.ref} -> 新: {new_range}")
    logger.debug(f"DEBUG: 原始表格有表頭: {original_has_header}, render_header: {render_header}")
    logger.debug(f"DEBUG: table_start_row: {table_start_row}, total_rows: {total_rows}")

    # 使用統一的表格屬性更新函式，確保 headerRowCount 正確設置
    include_header = render_header
    table_name = getattr(table_obj, 'name', 'Unknown')

    update_table_range_sync(table_obj, new_range, f"[Block渲染]{table_name}", dataframe, include_header)
