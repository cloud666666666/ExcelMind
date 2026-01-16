"""Excel 操作工具集 - 聚合所有技能工具"""

from typing import Any, Callable, Dict, List

# 从各技能模块导入工具
from .skills.core_query.tools import TOOLS as CORE_QUERY_TOOLS
from .skills.aggregation.tools import TOOLS as AGGREGATION_TOOLS
from .skills.visualization.tools import TOOLS as VISUALIZATION_TOOLS
from .skills.modification.tools import TOOLS as MODIFICATION_TOOLS
from .skills.formula.tools import TOOLS as FORMULA_TOOLS
from .skills.formatting.tools import TOOLS as FORMATTING_TOOLS
from .skills.sheet_management.tools import TOOLS as SHEET_MANAGEMENT_TOOLS
from .skills.utility.tools import TOOLS as UTILITY_TOOLS
from .skills.calculation.tools import TOOLS as CALCULATION_TOOLS

# 重新导出各技能的工具（向后兼容）
from .skills.core_query.tools import (
    filter_data,
    search_data,
    get_data_preview,
    get_column_stats,
    get_unique_values,
)
from .skills.aggregation.tools import (
    aggregate_data,
    group_and_aggregate,
    sort_data,
)
from .skills.visualization.tools import generate_chart
from .skills.modification.tools import (
    write_cell,
    write_range,
    insert_rows,
    delete_rows,
    save_file,
    save_to_original,
    export_file,
    quick_export,
    get_change_log,
)
from .skills.formula.tools import write_formula, read_formula
from .skills.formatting.tools import (
    set_font,
    set_fill,
    set_alignment,
    set_border,
    set_number_format,
    set_cell_style,
    merge_cells,
    unmerge_cells,
    set_column_width,
    set_row_height,
    auto_fit_column,
)
from .skills.sheet_management.tools import switch_sheet
from .skills.utility.tools import get_current_time
from .skills.calculation.tools import calculate

# 重新导出共享辅助函数（向后兼容）
from .skills._common import _limit_result, _df_to_result, _get_filter_mask


# ==================== 工具分类 ====================

# 查询工具
QUERY_TOOLS = CORE_QUERY_TOOLS

# 聚合工具
AGG_TOOLS = AGGREGATION_TOOLS

# 可视化工具
VIS_TOOLS = VISUALIZATION_TOOLS

# 写入工具
WRITE_TOOLS = MODIFICATION_TOOLS

# 公式工具
FORMULA_TOOLS_LIST = FORMULA_TOOLS

# 格式化工具
FORMAT_TOOLS = FORMATTING_TOOLS

# 系统工具
SYSTEM_TOOLS = SHEET_MANAGEMENT_TOOLS + UTILITY_TOOLS + CALCULATION_TOOLS


# ==================== 工具注册表 ====================

def get_tools_registry() -> Dict[str, Callable]:
    """获取所有工具的注册表（工具名 -> 工具函数）"""
    registry = {}
    for tool_list in [
        CORE_QUERY_TOOLS,
        AGGREGATION_TOOLS,
        VISUALIZATION_TOOLS,
        MODIFICATION_TOOLS,
        FORMULA_TOOLS,
        FORMATTING_TOOLS,
        SHEET_MANAGEMENT_TOOLS,
        UTILITY_TOOLS,
        CALCULATION_TOOLS,
    ]:
        for tool_func in tool_list:
            registry[tool_func.name] = tool_func
    return registry


def get_tools_by_names(tool_names: List[str]) -> List[Callable]:
    """根据工具名称列表获取工具函数列表"""
    registry = get_tools_registry()
    tools = []
    for name in tool_names:
        if name in registry:
            tools.append(registry[name])
    return tools


# ==================== 按技能分组的工具 ====================

SKILL_TOOLS = {
    "core_query": CORE_QUERY_TOOLS,
    "aggregation": AGGREGATION_TOOLS,
    "visualization": VISUALIZATION_TOOLS,
    "modification": MODIFICATION_TOOLS,
    "formula": FORMULA_TOOLS,
    "formatting": FORMATTING_TOOLS,
    "sheet_management": SHEET_MANAGEMENT_TOOLS,
    "utility": UTILITY_TOOLS,
    "calculation": CALCULATION_TOOLS,
}


# ==================== 导出所有工具列表 ====================

ALL_TOOLS = (
    CORE_QUERY_TOOLS +
    AGGREGATION_TOOLS +
    VISUALIZATION_TOOLS +
    MODIFICATION_TOOLS +
    FORMULA_TOOLS +
    FORMATTING_TOOLS +
    SHEET_MANAGEMENT_TOOLS +
    UTILITY_TOOLS +
    CALCULATION_TOOLS
)
