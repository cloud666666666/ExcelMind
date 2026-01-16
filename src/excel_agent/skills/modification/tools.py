"""数据修改工具"""

from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Optional
import os

from langchain_core.tools import tool

from ...excel_loader import get_loader


@tool
def write_cell(
    cell: str,
    value: Any,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """写入单个单元格的值。

    Args:
        cell: 单元格地址，如 "A1", "B2", "C10"
        value: 要写入的值
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
        result = loader.write_cell(cell, value, sheet=sheet)
        return {
            "success": True,
            "cell": cell,
            "old_value": result.get("old_value"),
            "new_value": result.get("new_value"),
            "message": f"已将 {cell} 的值从 '{result.get('old_value')}' 修改为 '{value}'"
        }
    except Exception as e:
        return {"error": f"写入失败: {str(e)}"}


@tool
def write_range(
    start_cell: str,
    data: List[List[Any]],
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """批量写入数据到指定范围。

    Args:
        start_cell: 起始单元格地址，如 "A1"
        data: 二维数据数组，如 [[1, 2, 3], [4, 5, 6]]
        sheet: 目标工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持写入操作（需要双引擎模式）"}
    if not data or not isinstance(data, list):
        return {"error": "data 必须是非空的二维数组"}

    try:
        result = loader.write_range(start_cell, data, sheet=sheet)
        return {
            "success": True,
            "start_cell": start_cell,
            "end_cell": result.get("end_cell"),
            "rows_written": result.get("rows_written"),
            "cells_written": result.get("cells_written"),
            "message": f"已从 {start_cell} 开始写入 {result.get('rows_written')} 行数据"
        }
    except Exception as e:
        return {"error": f"批量写入失败: {str(e)}"}


@tool
def insert_rows(
    row: int,
    count: int = 1,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """在指定位置插入行。

    Args:
        row: 在此行之前插入（行号从 1 开始）
        count: 插入的行数，默认 1
        sheet: 工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持写入操作（需要双引擎模式）"}
    if row < 1:
        return {"error": "行号必须大于等于 1"}
    if count < 1:
        return {"error": "插入行数必须大于等于 1"}

    try:
        loader.insert_rows(row, count, sheet=sheet)
        return {
            "success": True,
            "row": row,
            "count": count,
            "message": f"已在第 {row} 行之前插入 {count} 行"
        }
    except Exception as e:
        return {"error": f"插入行失败: {str(e)}"}


@tool
def delete_rows(
    start_row: int,
    end_row: Optional[int] = None,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """删除指定范围的行。

    Args:
        start_row: 起始行号（从 1 开始）
        end_row: 结束行号（含），默认只删除起始行
        sheet: 工作表名称，默认当前工作表

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持写入操作（需要双引擎模式）"}
    if start_row < 1:
        return {"error": "行号必须大于等于 1"}
    if end_row is not None and end_row < start_row:
        return {"error": "结束行号不能小于起始行号"}

    try:
        result = loader.delete_rows(start_row, end_row, sheet=sheet)
        count = result.get("count", 1)
        if end_row and end_row != start_row:
            message = f"已删除第 {start_row} 行到第 {end_row} 行，共 {count} 行"
        else:
            message = f"已删除第 {start_row} 行"
        return {
            "success": True,
            "start_row": start_row,
            "end_row": end_row or start_row,
            "count": count,
            "message": message
        }
    except Exception as e:
        return {"error": f"删除行失败: {str(e)}"}


@tool
def save_file(path: Optional[str] = None) -> Dict[str, Any]:
    """保存当前 Excel 文件到副本。

    Args:
        path: 保存路径，默认保存到副本文件

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持保存操作（需要双引擎模式）"}

    try:
        result = loader.save_table(file_path=path)
        save_path = result.get("file_path")
        message = f"文件已另存为: {save_path}" if path else f"文件已保存到副本: {save_path}"
        return {
            "success": True,
            "file_path": save_path,
            "message": message,
            "note": "修改已保存到工作副本，原始文件未受影响"
        }
    except Exception as e:
        return {"error": f"保存失败: {str(e)}"}


@tool
def save_to_original() -> Dict[str, Any]:
    """将修改保存回原始文件。警告：此操作会覆盖原始文件！

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持保存操作（需要双引擎模式）"}

    try:
        result = loader.save_to_original()
        return {
            "success": True,
            "original_path": result.get("original_path"),
            "message": result.get("message"),
            "warning": "原始文件已被覆盖"
        }
    except Exception as e:
        return {"error": f"保存到原始文件失败: {str(e)}"}


@tool
def export_file(export_path: str) -> Dict[str, Any]:
    """导出当前文件到新位置。

    Args:
        export_path: 导出的目标路径（完整路径，包含文件名）

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持导出操作（需要双引擎模式）"}

    try:
        export_path_obj = Path(export_path).expanduser().absolute()
        parent_dir = export_path_obj.parent

        if not parent_dir.exists():
            try:
                parent_dir.mkdir(parents=True, exist_ok=True)
            except Exception as dir_error:
                return {"error": f"无法创建目录 {parent_dir}: {str(dir_error)}"}

        if not os.access(str(parent_dir), os.W_OK):
            return {"error": f"没有权限写入目录: {parent_dir}"}

        if export_path_obj.suffix.lower() not in ['.xlsx', '.xlsm']:
            export_path_obj = export_path_obj.with_suffix('.xlsx')

        result = loader.export_to(str(export_path_obj))
        actual_path = result.get("export_path", str(export_path_obj))

        if not Path(actual_path).exists():
            return {"error": f"导出似乎成功但文件未创建: {actual_path}"}

        return {
            "success": True,
            "export_path": actual_path,
            "message": f"文件已成功导出到: {actual_path}",
            "file_exists": True
        }
    except Exception as e:
        return {"error": f"导出失败: {str(e)}"}


@tool
def quick_export(filename_suffix: str = "modified") -> Dict[str, Any]:
    """快速导出文件到原文件所在目录或用户下载目录。

    Args:
        filename_suffix: 文件名后缀，默认为 "modified"

    Returns:
        操作结果
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持导出操作（需要双引擎模式）"}

    try:
        table_info = loader.get_active_table_info()
        if not table_info:
            return {"error": "无法获取当前表信息"}

        original_path = Path(table_info.original_path)
        original_filename = table_info.filename
        stem = Path(original_filename).stem
        suffix = Path(original_filename).suffix or '.xlsx'

        is_temp_file = 'temp' in str(original_path).lower() or 'tmp' in str(original_path).lower()

        if is_temp_file:
            downloads_dir = Path.home() / "Downloads"
            if not downloads_dir.exists():
                downloads_dir = Path.home()
            parent_dir = downloads_dir
        else:
            parent_dir = original_path.parent

        new_filename = f"{stem}_{filename_suffix}{suffix}"
        export_path = parent_dir / new_filename

        if export_path.exists():
            timestamp = datetime.now().strftime("%H%M%S")
            new_filename = f"{stem}_{filename_suffix}_{timestamp}{suffix}"
            export_path = parent_dir / new_filename

        result = loader.export_to(str(export_path))
        actual_path = result.get("export_path", str(export_path))

        if not Path(actual_path).exists():
            return {"error": f"导出失败，文件未创建: {actual_path}"}

        return {
            "success": True,
            "export_path": actual_path,
            "directory": str(parent_dir),
            "filename": new_filename,
            "message": f"文件已导出到: {actual_path}",
            "tip": f"请在文件夹 {parent_dir} 中查找文件 {new_filename}"
        }
    except Exception as e:
        return {"error": f"导出失败: {str(e)}"}


@tool
def get_change_log() -> Dict[str, Any]:
    """获取当前文件的变更记录。

    Returns:
        变更记录列表
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}
    if not loader.is_dual_engine():
        return {"error": "当前表不支持此操作（需要双引擎模式）"}

    try:
        changes = loader.get_change_log()
        return {
            "total_changes": len(changes),
            "changes": changes,
            "has_unsaved_changes": len(changes) > 0,
            "message": f"共有 {len(changes)} 条变更记录" if changes else "没有变更记录"
        }
    except Exception as e:
        return {"error": f"获取变更记录失败: {str(e)}"}


# 导出工具列表
TOOLS = [
    write_cell,
    write_range,
    insert_rows,
    delete_rows,
    save_file,
    save_to_original,
    export_file,
    quick_export,
    get_change_log,
]
