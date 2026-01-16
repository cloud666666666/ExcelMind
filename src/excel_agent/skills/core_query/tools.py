"""数据查询工具"""

from typing import Any, Dict, List, Optional
import pandas as pd
from langchain_core.tools import tool

from .._common import _df_to_result, _get_filter_mask
from ...excel_loader import get_loader


@tool
def filter_data(
    column: Optional[str] = None,
    operator: Optional[str] = None,
    value: Optional[Any] = None,
    filters: Optional[List[Dict[str, Any]]] = None,
    select_columns: Optional[List[str]] = None,
    sort_by: Optional[str] = None,
    ascending: bool = True,
    limit: int = 20
) -> Dict[str, Any]:
    """按条件筛选 Excel 数据，支持排序。

    Args:
        column: 单条件筛选时的列名
        operator: 比较运算符: ==, !=, >, <, >=, <=, contains, startswith, endswith
        value: 比较值
        filters: 多条件筛选列表
        select_columns: 指定返回的列名列表
        sort_by: 排序列名
        ascending: 排序方向
        limit: 返回结果数量限制
    """
    loader = get_loader()
    df = loader.dataframe.copy()

    try:
        final_mask = pd.Series([True] * len(df))

        if column and operator and value is not None:
            mask = _get_filter_mask(df, column, operator, value)
            final_mask &= mask

        if filters:
            for f in filters:
                f_col, f_op, f_val = f.get("column"), f.get("operator"), f.get("value")
                if f_col and f_op and f_val is not None:
                    final_mask &= _get_filter_mask(df, f_col, f_op, f_val)

        result_df = df[final_mask]

        if sort_by:
            if sort_by not in result_df.columns:
                return {"error": f"排序列 '{sort_by}' 不存在"}
            result_df = result_df.sort_values(by=sort_by, ascending=ascending)

        return _df_to_result(result_df, limit, select_columns)
    except Exception as e:
        return {"error": f"筛选出错: {str(e)}"}


@tool
def search_data(
    keyword: str,
    columns: Optional[List[str]] = None,
    select_columns: Optional[List[str]] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """在指定列或所有列中搜索关键词。"""
    loader = get_loader()
    df = loader.dataframe

    try:
        search_cols = columns if columns else df.columns
        mask = pd.Series([False] * len(df))
        for col in search_cols:
            if col in df.columns:
                mask |= df[col].astype(str).str.contains(keyword, case=False, na=False)

        return _df_to_result(df[mask], limit, select_columns)
    except Exception as e:
        return {"error": f"搜索出错: {str(e)}"}


@tool
def get_data_preview(n_rows: int = 10) -> Dict[str, Any]:
    """获取数据预览。"""
    loader = get_loader()
    active_loader = loader.get_active_loader()
    if active_loader is None:
        return {"error": "没有活跃的表"}
    return active_loader.get_preview(n_rows)


@tool
def get_column_stats(
    column: str,
    filters: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """获取指定列的详细统计信息。"""
    loader = get_loader()
    df = loader.dataframe.copy()

    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col, f_op, f_val = f.get("column"), f.get("operator"), f.get("value")
                if f_col and f_op and f_val is not None:
                    final_mask &= _get_filter_mask(df, f_col, f_op, f_val)
            df = df[final_mask]
        except Exception as e:
            return {"error": f"筛选条件错误: {str(e)}"}

    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在"}

    col = df[column]
    try:
        stats = {
            "column": column,
            "filtered_rows": len(df),
            "dtype": str(col.dtype),
            "count": int(col.count()),
            "null_count": int(col.isna().sum()),
            "unique_count": int(col.nunique()),
        }

        if pd.api.types.is_numeric_dtype(col):
            stats.update({
                "min": float(col.min()) if not col.isna().all() else None,
                "max": float(col.max()) if not col.isna().all() else None,
                "mean": float(col.mean()) if not col.isna().all() else None,
                "median": float(col.median()) if not col.isna().all() else None,
            })

        return stats
    except Exception as e:
        return {"error": f"统计出错: {str(e)}"}


@tool
def get_unique_values(
    column: str,
    filters: Optional[List[Dict[str, Any]]] = None,
    limit: int = 50
) -> Dict[str, Any]:
    """获取指定列的唯一值列表。"""
    loader = get_loader()
    df = loader.dataframe.copy()

    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col, f_op, f_val = f.get("column"), f.get("operator"), f.get("value")
                if f_col and f_op and f_val is not None:
                    final_mask &= _get_filter_mask(df, f_col, f_op, f_val)
            df = df[final_mask]
        except Exception as e:
            return {"error": f"筛选条件错误: {str(e)}"}

    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在"}

    try:
        value_counts = df[column].value_counts()
        total_unique = len(value_counts)
        if limit:
            value_counts = value_counts.head(limit)

        values = [{"value": str(idx), "count": int(count)} for idx, count in value_counts.items()]

        return {
            "column": column,
            "filtered_rows": len(df),
            "total_unique": total_unique,
            "returned_unique": len(values),
            "values": values,
        }
    except Exception as e:
        return {"error": f"获取唯一值出错: {str(e)}"}


# 导出工具列表
TOOLS = [filter_data, search_data, get_data_preview, get_column_stats, get_unique_values]
