"""工作表管理工具"""

from typing import Any, Dict
from langchain_core.tools import tool

from ...excel_loader import get_loader


@tool
def switch_sheet(sheet_name: str) -> Dict[str, Any]:
    """切换当前活跃表的工作表（Sheet）。

    Args:
        sheet_name: 要切换到的工作表名称

    Returns:
        切换后的工作表结构信息
    """
    loader = get_loader()
    active_loader = loader.get_active_loader()

    if active_loader is None:
        return {"error": "没有活跃的表"}

    try:
        structure = active_loader.switch_sheet(sheet_name)

        active_info = loader.get_active_table_info()
        if active_info:
            active_info.sheet_name = structure["sheet_name"]
            active_info.total_rows = structure["total_rows"]
            active_info.total_columns = structure["total_columns"]

        return {
            "message": f"已切换到工作表: {sheet_name}",
            "sheet_name": structure["sheet_name"],
            "all_sheets": structure["all_sheets"],
            "total_rows": structure["total_rows"],
            "total_columns": structure["total_columns"],
            "columns": [col["name"] for col in structure["columns"]],
        }
    except ValueError as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"切换工作表失败: {str(e)}"}


# 导出工具列表
TOOLS = [switch_sheet]
