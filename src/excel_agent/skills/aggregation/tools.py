"""聚合分析工具"""

from typing import Any, Dict, List, Optional
import pandas as pd
from langchain_core.tools import tool

from .._common import _df_to_result, _get_filter_mask
from ...excel_loader import get_loader


@tool
def aggregate_data(
    column: str,
    agg_func: str,
    filters: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """对指定列进行聚合统计。可选先筛选数据再聚合。

    Args:
        column: 要统计的列名
        agg_func: 聚合函数，可选值: sum, mean, count, min, max, median, std
        filters: 可选的筛选条件列表

    Returns:
        统计结果
    """
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
        return {"error": f"列 '{column}' 不存在，可用列: {list(df.columns)}"}

    col = df[column]

    try:
        agg_funcs = {
            "sum": col.sum,
            "mean": col.mean,
            "count": col.count,
            "min": col.min,
            "max": col.max,
            "median": col.median,
            "std": col.std,
        }

        if agg_func not in agg_funcs:
            return {"error": f"不支持的聚合函数: {agg_func}"}

        result = agg_funcs[agg_func]()

        if hasattr(result, 'item'):
            result = result.item()

        return {
            "column": column,
            "function": agg_func,
            "filtered_rows": len(df),
            "result": result,
        }
    except Exception as e:
        return {"error": f"聚合计算出错: {str(e)}"}


@tool
def group_and_aggregate(
    group_by: str,
    agg_column: str,
    agg_func: str,
    filters: Optional[List[Dict[str, Any]]] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """按列分组并进行聚合统计。可选先筛选数据再分组。

    Args:
        group_by: 分组列名
        agg_column: 要聚合的列名
        agg_func: 聚合函数，可选值: sum, mean, count, min, max
        filters: 可选的筛选条件列表
        limit: 返回结果数量限制，默认20

    Returns:
        分组聚合结果
    """
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

    if group_by not in df.columns:
        return {"error": f"分组列 '{group_by}' 不存在，可用列: {list(df.columns)}"}
    if agg_column not in df.columns:
        return {"error": f"聚合列 '{agg_column}' 不存在，可用列: {list(df.columns)}"}

    try:
        grouped = df.groupby(group_by)[agg_column].agg(agg_func).reset_index()
        grouped.columns = [group_by, f"{agg_column}_{agg_func}"]
        grouped = grouped.sort_values(by=grouped.columns[1], ascending=False)

        result = _df_to_result(grouped, limit)
        result["filtered_rows"] = len(df)
        return result
    except Exception as e:
        return {"error": f"分组聚合出错: {str(e)}"}


@tool
def sort_data(
    column: str,
    ascending: bool = True,
    filters: Optional[List[Dict[str, Any]]] = None,
    select_columns: Optional[List[str]] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """按指定列排序数据。可选先筛选、指定返回列。

    Args:
        column: 排序列名
        ascending: 是否升序排列，默认True
        filters: 可选的筛选条件列表
        select_columns: 指定返回的列名列表
        limit: 返回结果数量限制，默认20

    Returns:
        排序后的数据
    """
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
        return {"error": f"列 '{column}' 不存在，可用列: {list(df.columns)}"}

    try:
        sorted_df = df.sort_values(by=column, ascending=ascending)
        return _df_to_result(sorted_df, limit, select_columns)
    except Exception as e:
        return {"error": f"排序出错: {str(e)}"}


# 导出工具列表
TOOLS = [aggregate_data, group_and_aggregate, sort_data]
