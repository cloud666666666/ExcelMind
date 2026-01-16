"""格式设置工具"""

from typing import Any, Dict, Optional
from langchain_core.tools import tool

from ...excel_loader import get_loader


@tool
def set_font(
    cell_range: str,
    name: Optional[str] = None,
    size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[str] = None,
    color: Optional[str] = None,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置单元格字体样式。

    Args:
        cell_range: 单元格或范围，如 "A1" 或 "A1:C10"
        name: 字体名称，如 "Arial", "宋体", "微软雅黑"
        size: 字号，如 12, 14, 16
        bold: 是否加粗
        italic: 是否斜体
        underline: 下划线类型 ("single", "double")
        color: 字体颜色 (十六进制，如 "FF0000" 红色)
        sheet: 工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_font(
            cell_range=cell_range, name=name, size=size, bold=bold,
            italic=italic, underline=underline, color=color, sheet=sheet
        )

        if not result.get("success"):
            return {"error": result.get("error", "设置字体失败")}

        settings = []
        if name: settings.append(f"字体: {name}")
        if size: settings.append(f"字号: {size}")
        if bold: settings.append("加粗")
        if italic: settings.append("斜体")
        if underline: settings.append(f"下划线: {underline}")
        if color: settings.append(f"颜色: #{color}")

        return {
            "success": True,
            "cell_range": cell_range,
            "cells_modified": result.get("cells_modified"),
            "message": f"已设置 {cell_range} 的字体: {', '.join(settings)}"
        }
    except Exception as e:
        return {"error": f"设置字体失败: {str(e)}"}


@tool
def set_fill(
    cell_range: str,
    color: str,
    fill_type: str = "solid",
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置单元格背景色。

    Args:
        cell_range: 单元格或范围，如 "A1" 或 "A1:C10"
        color: 背景颜色 (十六进制，如 "FFFF00" 黄色)
        fill_type: 填充类型，默认 "solid"
        sheet: 工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_fill(cell_range=cell_range, color=color, fill_type=fill_type, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置背景色失败")}

        return {
            "success": True,
            "cell_range": cell_range,
            "cells_modified": result.get("cells_modified"),
            "message": f"已设置 {cell_range} 的背景色为 #{color}"
        }
    except Exception as e:
        return {"error": f"设置背景色失败: {str(e)}"}


@tool
def set_alignment(
    cell_range: str,
    horizontal: Optional[str] = None,
    vertical: Optional[str] = None,
    wrap_text: Optional[bool] = None,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置单元格对齐方式。

    Args:
        cell_range: 单元格或范围，如 "A1" 或 "A1:C10"
        horizontal: 水平对齐 ("left", "center", "right")
        vertical: 垂直对齐 ("top", "center", "bottom")
        wrap_text: 是否自动换行
        sheet: 工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_alignment(
            cell_range=cell_range, horizontal=horizontal, vertical=vertical,
            wrap_text=wrap_text, sheet=sheet
        )

        if not result.get("success"):
            return {"error": result.get("error", "设置对齐失败")}

        settings = []
        if horizontal: settings.append(f"水平: {horizontal}")
        if vertical: settings.append(f"垂直: {vertical}")
        if wrap_text: settings.append("自动换行")

        return {
            "success": True,
            "cell_range": cell_range,
            "cells_modified": result.get("cells_modified"),
            "message": f"已设置 {cell_range} 的对齐方式: {', '.join(settings)}"
        }
    except Exception as e:
        return {"error": f"设置对齐失败: {str(e)}"}


@tool
def set_border(
    cell_range: str,
    style: str = "thin",
    color: str = "000000",
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置单元格边框。

    Args:
        cell_range: 单元格或范围，如 "A1" 或 "A1:C10"
        style: 边框样式 ("thin", "medium", "thick", "double", "dashed")
        color: 边框颜色 (十六进制，如 "000000" 黑色)
        sheet: 工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_border(cell_range=cell_range, style=style, color=color, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置边框失败")}

        return {
            "success": True,
            "cell_range": cell_range,
            "cells_modified": result.get("cells_modified"),
            "message": f"已为 {cell_range} 添加 {style} 边框"
        }
    except Exception as e:
        return {"error": f"设置边框失败: {str(e)}"}


@tool
def set_number_format(
    cell_range: str,
    format_code: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置单元格数字格式。

    常用格式代码:
    - "#,##0" - 千分位整数
    - "#,##0.00" - 千分位两位小数
    - "0.00%" - 百分比
    - "yyyy-mm-dd" - 日期
    - "¥#,##0.00" - 人民币
    - "$#,##0.00" - 美元

    Args:
        cell_range: 单元格或范围
        format_code: 数字格式代码
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_number_format(cell_range=cell_range, format_code=format_code, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置数字格式失败")}

        return {
            "success": True,
            "cell_range": cell_range,
            "cells_modified": result.get("cells_modified"),
            "format_code": format_code,
            "message": f"已设置 {cell_range} 的数字格式为: {format_code}"
        }
    except Exception as e:
        return {"error": f"设置数字格式失败: {str(e)}"}


@tool
def set_cell_style(
    cell_range: str,
    font_bold: Optional[bool] = None,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    horizontal: Optional[str] = None,
    border_style: Optional[str] = None,
    number_format: Optional[str] = None,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """一次性设置多种单元格样式。

    Args:
        cell_range: 单元格或范围
        font_bold: 是否加粗
        font_size: 字号
        font_color: 字体颜色
        bg_color: 背景颜色
        horizontal: 水平对齐
        border_style: 边框样式
        number_format: 数字格式代码
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_cell_style(
            cell_range=cell_range, font_bold=font_bold, font_size=font_size,
            font_color=font_color, bg_color=bg_color, horizontal=horizontal,
            border_style=border_style, number_format=number_format, sheet=sheet
        )

        if not result.get("success"):
            return {"error": result.get("error", "设置样式失败")}

        return {
            "success": True,
            "cell_range": cell_range,
            "cells_modified": result.get("cells_modified"),
            "settings_applied": result.get("settings_applied"),
            "message": f"已设置 {cell_range} 的样式: {', '.join(result.get('settings_applied', []))}"
        }
    except Exception as e:
        return {"error": f"设置样式失败: {str(e)}"}


@tool
def merge_cells(cell_range: str, sheet: Optional[str] = None) -> Dict[str, Any]:
    """合并单元格。

    Args:
        cell_range: 要合并的范围，如 "A1:C3"
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}
    if ":" not in cell_range:
        return {"error": "合并单元格需要指定范围，如 'A1:C3'"}

    try:
        doc = loader.get_active_document()
        result = doc.merge_cells(cell_range=cell_range, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "合并单元格失败")}

        return {"success": True, "cell_range": cell_range, "message": f"已合并单元格: {cell_range}"}
    except Exception as e:
        return {"error": f"合并单元格失败: {str(e)}"}


@tool
def unmerge_cells(cell_range: str, sheet: Optional[str] = None) -> Dict[str, Any]:
    """取消合并单元格。

    Args:
        cell_range: 要取消合并的范围
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.unmerge_cells(cell_range=cell_range, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "取消合并失败")}

        return {"success": True, "cell_range": cell_range, "message": f"已取消合并: {cell_range}"}
    except Exception as e:
        return {"error": f"取消合并失败: {str(e)}"}


@tool
def set_column_width(column: str, width: float, sheet: Optional[str] = None) -> Dict[str, Any]:
    """设置列宽。

    Args:
        column: 列标识，如 "A", "B", "C"
        width: 列宽（字符数），如 15, 20, 30
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_column_width(column=column, width=width, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置列宽失败")}

        return {
            "success": True, "column": column.upper(), "width": width,
            "message": f"已设置列 {column.upper()} 的宽度为 {width}"
        }
    except Exception as e:
        return {"error": f"设置列宽失败: {str(e)}"}


@tool
def set_row_height(row: int, height: float, sheet: Optional[str] = None) -> Dict[str, Any]:
    """设置行高。

    Args:
        row: 行号（从 1 开始）
        height: 行高（磅），如 20, 30, 40
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.set_row_height(row=row, height=height, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置行高失败")}

        return {
            "success": True, "row": row, "height": height,
            "message": f"已设置第 {row} 行的高度为 {height} 磅"
        }
    except Exception as e:
        return {"error": f"设置行高失败: {str(e)}"}


@tool
def auto_fit_column(column: str, sheet: Optional[str] = None) -> Dict[str, Any]:
    """自动调整列宽以适应内容。

    Args:
        column: 列标识，如 "A", "B", "C"
        sheet: 工作表名称

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持格式化操作（需要双引擎模式）"}

    try:
        doc = loader.get_active_document()
        result = doc.auto_fit_column(column=column, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "自动调整列宽失败")}

        return {
            "success": True, "column": column.upper(), "width": result.get("width"),
            "message": f"已自动调整列 {column.upper()} 的宽度为 {result.get('width')}"
        }
    except Exception as e:
        return {"error": f"自动调整列宽失败: {str(e)}"}


# 导出工具列表
TOOLS = [
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
]
