"""Excel 公式工具"""

from typing import Any, Dict, Optional
from langchain_core.tools import tool

from ...excel_loader import get_loader


@tool
def write_formula(
    cell: str,
    formula: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """写入 Excel 公式到指定单元格。

    注意：公式不会立即计算，需要用 Excel 打开文件后才会计算结果。

    Args:
        cell: 目标单元格地址，如 "C1"
        formula: Excel 公式，如 "SUM(A1:B1)" 或 "=SUM(A1:B1)"
        sheet: 目标工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持写入操作（需要双引擎模式）"}

    try:
        result = loader.write_formula(cell, formula, sheet=sheet)
        return {
            "success": True,
            "cell": cell,
            "formula": result.get("formula"),
            "message": f"已在 {cell} 写入公式: {result.get('formula')}",
            "note": "公式将在 Excel 中打开时计算"
        }
    except Exception as e:
        return {"error": f"写入公式失败: {str(e)}"}


@tool
def read_formula(
    cell: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """读取单元格的公式。

    Args:
        cell: 单元格地址，如 "A1"
        sheet: 工作表名称，默认当前工作表

    Returns:
        公式内容，如果不是公式则返回 None
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持此操作（需要双引擎模式）"}

    try:
        formula = loader.read_formula(cell, sheet=sheet)
        if formula:
            return {
                "cell": cell,
                "formula": f"={formula}",
                "has_formula": True
            }
        else:
            doc = loader.get_active_document()
            value = doc.read_cell(cell, sheet) if doc else None
            return {
                "cell": cell,
                "formula": None,
                "has_formula": False,
                "value": value,
                "message": f"单元格 {cell} 不包含公式，当前值为: {value}"
            }
    except Exception as e:
        return {"error": f"读取公式失败: {str(e)}"}


# 导出工具列表
TOOLS = [write_formula, read_formula]
