"""共享辅助函数 - 供各技能工具使用"""

from typing import Any, Dict, List, Optional
import pandas as pd

from ..config import get_config
from ..excel_loader import get_loader


def _limit_result(df: pd.DataFrame, limit: Optional[int] = None) -> pd.DataFrame:
    """限制返回结果行数"""
    config = get_config()
    if limit is None:
        limit = config.excel.default_result_limit
    limit = min(limit, config.excel.max_result_limit)
    return df.head(limit)


def _df_to_result(df: pd.DataFrame, limit: Optional[int] = None, select_columns: Optional[List[str]] = None) -> Dict[str, Any]:
    """将 DataFrame 转换为结果字典"""
    if select_columns:
        available_cols = [c for c in select_columns if c in df.columns]
        if available_cols:
            df = df[available_cols]

    limited_df = _limit_result(df, limit)
    return {
        "total_rows": len(df),
        "returned_rows": len(limited_df),
        "columns": list(limited_df.columns),
        "data": limited_df.to_dict(orient="records"),
    }


def _get_filter_mask(df: pd.DataFrame, column: str, operator: str, value: Any) -> pd.Series:
    """生成单个筛选条件的布尔掩码"""
    if column not in df.columns:
        raise ValueError(f"列 '{column}' 不存在，可用列: {list(df.columns)}")

    col = df[column]

    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        numeric_value = None

    compare_value = numeric_value if numeric_value is not None else value

    if operator == "==":
        return col == compare_value
    elif operator == "!=":
        return col != compare_value
    elif operator == ">":
        return col > compare_value
    elif operator == "<":
        return col < compare_value
    elif operator == ">=":
        return col >= compare_value
    elif operator == "<=":
        return col <= compare_value
    elif operator == "contains":
        return col.astype(str).str.contains(str(value), case=False, na=False)
    elif operator == "startswith":
        return col.astype(str).str.startswith(str(value), na=False)
    elif operator == "endswith":
        return col.astype(str).str.endswith(str(value), na=False)
    else:
        raise ValueError(f"不支持的运算符: {operator}")
