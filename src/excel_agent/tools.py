"""Excel 操作工具集"""

from typing import Any, Dict, List, Optional

import pandas as pd
from langchain_core.tools import tool

from .excel_loader import get_loader
from .config import get_config


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
        # 确保请求的列存在
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
    """内部辅助函数：生成单个筛选条件的布尔掩码"""
    if column not in df.columns:
        raise ValueError(f"列 '{column}' 不存在，可用列: {list(df.columns)}")
    
    col = df[column]
    
    # 尝试将 value 转换为数值进行比较
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
        operator: 比较运算符，仅支持: ==, !=, >, <, >=, <=, contains, startswith, endswith
                  注意: 不支持 between/equals 等运算符，请用多个 >= 和 <= 条件代替 between
        value: 单条件筛选时的比较值（支持字符串、数值等任意类型）
        filters: 多条件筛选列表，每个元素为 {"column": "...", "operator": "...", "value": ...}
                 operator 同样仅支持上述列出的运算符
        select_columns: 指定返回的列名列表，为空则返回所有列
        sort_by: 排序列名，可选
        ascending: 排序方向，True为升序，False为降序，默认True
        limit: 返回结果数量限制，默认20
        
    Returns:
        筛选后的数据（可选排序）
    """
    loader = get_loader()
    df = loader.dataframe.copy()
    
    try:
        # 初始掩码为全 True
        final_mask = pd.Series([True] * len(df))
        
        # 1. 处理单条件参数 (兼容旧调用)
        if column and operator and value is not None:
            mask = _get_filter_mask(df, column, operator, value)
            final_mask &= mask
            
        # 2. 处理多条件列表
        if filters:
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
        
        result_df = df[final_mask]
        
        # 3. 排序（如果指定了 sort_by）
        if sort_by:
            if sort_by not in result_df.columns:
                return {"error": f"排序列 '{sort_by}' 不存在，可用列: {list(result_df.columns)}"}
            result_df = result_df.sort_values(by=sort_by, ascending=ascending)
        
        return _df_to_result(result_df, limit, select_columns)
    except Exception as e:
        return {"error": f"筛选出错: {str(e)}"}


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
        filters: 可选的筛选条件列表，每个元素为 {"column": "...", "operator": "...", "value": ...}
                 operator 仅支持: ==, !=, >, <, >=, <=, contains, startswith, endswith
                 注意: 不支持 between/equals，用 >= 和 <= 组合代替
        
    Returns:
        统计结果
    """
    loader = get_loader()
    df = loader.dataframe.copy()
    
    # 如果有筛选条件，先进行筛选
    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
            df = df[final_mask]
        except Exception as e:
            return {"error": f"筛选条件错误: {str(e)}"}
    
    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在，可用列: {list(df.columns)}"}
    
    col = df[column]
    
    try:
        if agg_func == "sum":
            result = col.sum()
        elif agg_func == "mean":
            result = col.mean()
        elif agg_func == "count":
            result = col.count()
        elif agg_func == "min":
            result = col.min()
        elif agg_func == "max":
            result = col.max()
        elif agg_func == "median":
            result = col.median()
        elif agg_func == "std":
            result = col.std()
        else:
            return {"error": f"不支持的聚合函数: {agg_func}"}
        
        # 处理 numpy 类型
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
        filters: 可选的筛选条件列表，operator 仅支持: ==, !=, >, <, >=, <=, contains, startswith, endswith
        limit: 返回结果数量限制，默认20
        
    Returns:
        分组聚合结果
    """
    loader = get_loader()
    df = loader.dataframe.copy()
    
    # 如果有筛选条件，先进行筛选
    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
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
        
        # 按聚合结果降序排序
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
    
    # 如果有筛选条件，先进行筛选
    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
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


@tool
def search_data(
    keyword: str, 
    columns: Optional[List[str]] = None,
    select_columns: Optional[List[str]] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """在指定列或所有列中搜索关键词。
    
    Args:
        keyword: 搜索关键词
        columns: 要搜索的列名列表，为空则搜索所有列
        select_columns: 指定返回的列名列表
        limit: 返回结果数量限制，默认20
        
    Returns:
        包含关键词的数据行
    """
    loader = get_loader()
    df = loader.dataframe
    
    try:
        # 确定搜索范围
        search_cols = columns if columns else df.columns
        
        # 在指定列中搜索
        mask = pd.Series([False] * len(df))
        for col in search_cols:
            if col in df.columns:
                mask |= df[col].astype(str).str.contains(keyword, case=False, na=False)
        
        result_df = df[mask]
        return _df_to_result(result_df, limit, select_columns)
    except Exception as e:
        return {"error": f"搜索出错: {str(e)}"}


@tool
def get_column_stats(
    column: str,
    filters: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """获取指定列的详细统计信息。可选先筛选数据再统计。
    
    Args:
        column: 列名
        filters: 可选的筛选条件列表
        
    Returns:
        列的统计信息
    """
    loader = get_loader()
    df = loader.dataframe.copy()
    
    # 如果有筛选条件，先进行筛选
    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
            df = df[final_mask]
        except Exception as e:
            return {"error": f"筛选条件错误: {str(e)}"}
    
    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在，可用列: {list(df.columns)}"}
    
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
        
        # 数值类型额外统计
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
    """获取指定列的唯一值列表。可选先筛选数据。
    
    Args:
        column: 列名
        filters: 可选的筛选条件列表
        limit: 返回唯一值数量限制，默认50
        
    Returns:
        唯一值列表及其计数
    """
    loader = get_loader()
    df = loader.dataframe.copy()
    
    # 如果有筛选条件，先进行筛选
    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
            df = df[final_mask]
        except Exception as e:
            return {"error": f"筛选条件错误: {str(e)}"}
    
    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在，可用列: {list(df.columns)}"}
    
    try:
        value_counts = df[column].value_counts()
        total_unique = len(value_counts)
        
        if limit:
            value_counts = value_counts.head(limit)
        
        values = [
            {"value": str(idx), "count": int(count)}
            for idx, count in value_counts.items()
        ]
        
        return {
            "column": column,
            "filtered_rows": len(df),
            "total_unique": total_unique,
            "returned_unique": len(values),
            "values": values,
        }
    except Exception as e:
        return {"error": f"获取唯一值出错: {str(e)}"}


@tool
def get_data_preview(n_rows: int = 10) -> Dict[str, Any]:
    """获取数据预览。

    Args:
        n_rows: 预览行数，默认10行

    Returns:
        数据预览
    """
    loader = get_loader()
    active_loader = loader.get_active_loader()
    if active_loader is None:
        return {"error": "没有活跃的表"}
    return active_loader.get_preview(n_rows)


@tool
def switch_sheet(sheet_name: str) -> Dict[str, Any]:
    """切换当前活跃表的工作表（Sheet）。

    Args:
        sheet_name: 要切换到的工作表名称

    Returns:
        切换后的工作表结构信息，包含列信息和数据规模
    """
    loader = get_loader()
    active_loader = loader.get_active_loader()

    if active_loader is None:
        return {"error": "没有活跃的表"}

    try:
        structure = active_loader.switch_sheet(sheet_name)

        # 更新 MultiExcelLoader 中的表信息
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


@tool
def get_current_time() -> Dict[str, Any]:
    """获取当前系统时间。
    
    Returns:
        当前时间信息
    """
    from datetime import datetime
    now = datetime.now()
    return {
        "current_time": now.strftime("%Y-%m-%d %H:%M:%S"),
        "weekday": now.strftime("%A"),
        "timestamp": now.timestamp()
    }


@tool
def calculate(expressions: List[str]) -> Dict[str, Any]:
    """执行数学计算。
    
    Args:
        expressions: 数学表达式列表，例如 ["(100+200)*0.5", "500/2"]
        
    Returns:
        每个表达式的计算结果
    """
    import math
    
    results = {}
    
    # 定义安全的计算环境
    safe_env = {
        "abs": abs,
        "round": round,
        "min": min,
        "max": max,
        "sum": sum,
        "pow": pow,
        "math": math,
    }
    
    for expr in expressions:
        try:
            # 移除危险字符，防止恶意代码
            if any(char in expr for char in ["__", "import", "eval", "exec", "open"]):
                results[expr] = "Error: Unsafe expression"
                continue
                
            # 执行计算
            result = eval(expr, {"__builtins__": None}, safe_env)
            results[expr] = result
        except Exception as e:
            results[expr] = f"Error: {str(e)}"
            
    return {"results": results}


@tool
def generate_chart(
    chart_type: Optional[str] = None,
    x_column: Optional[str] = None,
    y_column: Optional[str] = None,
    agg_column: Optional[str] = None,  # y_column 的别名，用于饼图等分组聚合场景
    group_by: Optional[str] = None,
    agg_func: str = "sum",
    title: str = "",
    filters: Optional[List[Dict[str, Any]]] = None,
    series_columns: Optional[List[str]] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """生成 ECharts 可视化图表配置。
    
    Args:
        chart_type: 图表类型，可选: bar(柱状图), line(折线图), pie(饼图), 
                   scatter(散点图), radar(雷达图), funnel(漏斗图)。
                   为空或"auto"时自动推荐。
        x_column: X轴数据列名（分类轴）
        y_column: Y轴数据列名（数值轴，单系列时使用）
        agg_column: 聚合列名（y_column 的别名，用于饼图等场景）
        group_by: 分组列名（用于饼图和多系列图）
        agg_func: 聚合函数: sum, mean, count, min, max
        title: 图表标题
        filters: 筛选条件列表，operator 仅支持: ==, !=, >, <, >=, <=, contains, startswith, endswith
                 注意: 不支持 between/equals，请用 >= 和 <= 组合代替 between
        series_columns: 多系列Y轴列名列表
        limit: 数据点数量限制，默认20
        
    Returns:
        包含 ECharts 配置的字典 {"chart": {...}, "message": "..."}
    """
    # 处理 agg_column 作为 y_column 的别名
    if agg_column and not y_column:
        y_column = agg_column
    
    loader = get_loader()
    df = loader.dataframe.copy()
    
    # 应用筛选条件
    if filters:
        try:
            final_mask = pd.Series([True] * len(df))
            for f in filters:
                f_col = f.get("column")
                f_op = f.get("operator")
                f_val = f.get("value")
                if f_col and f_op and f_val is not None:
                    mask = _get_filter_mask(df, f_col, f_op, f_val)
                    final_mask &= mask
            df = df[final_mask]
        except Exception as e:
            return {"error": f"筛选条件错误: {str(e)}"}
    
    if len(df) == 0:
        return {"error": "筛选后无数据，无法生成图表"}
    
    # 自动推荐图表类型
    def recommend_chart_type() -> str:
        """根据数据特征推荐图表类型"""
        if group_by and y_column:
            # 分组场景：检查分组数量
            unique_groups = df[group_by].nunique() if group_by in df.columns else 0
            if unique_groups <= 8:
                return "pie"  # 少量分组适合饼图
            return "bar"  # 多分组适合柱状图
        
        if x_column and y_column:
            x_dtype = df[x_column].dtype if x_column in df.columns else None
            y_dtype = df[y_column].dtype if y_column in df.columns else None
            
            # 两个数值列 → 散点图
            if pd.api.types.is_numeric_dtype(x_dtype) and pd.api.types.is_numeric_dtype(y_dtype):
                return "scatter"
            
            # X轴是日期/时间类型 → 折线图
            if pd.api.types.is_datetime64_any_dtype(x_dtype):
                return "line"
            
            # 默认柱状图
            return "bar"
        
        # 仅有分组列 → 饼图
        if group_by:
            return "pie"
        
        return "bar"
    
    # 确定图表类型
    final_chart_type = chart_type if chart_type and chart_type != "auto" else recommend_chart_type()
    
    try:
        # 准备图表数据
        chart_data = _prepare_chart_data(df, final_chart_type, x_column, y_column, 
                                         group_by, agg_func, series_columns, limit)
        
        if "error" in chart_data:
            return chart_data
        
        # 生成 ECharts 配置
        chart_config = _build_echart_config(final_chart_type, chart_data, title)
        
        chart_type_names = {
            "bar": "柱状图", "line": "折线图", "pie": "饼图",
            "scatter": "散点图", "radar": "雷达图", "funnel": "漏斗图"
        }
        message = f"已生成{chart_type_names.get(final_chart_type, final_chart_type)}，共 {chart_data.get('data_count', 0)} 个数据点。"
        
        return {
            "chart": chart_config,
            "chart_type": final_chart_type,
            "message": message
        }
    except Exception as e:
        return {"error": f"生成图表出错: {str(e)}"}


def _prepare_chart_data(df: pd.DataFrame, chart_type: str, x_column: Optional[str],
                        y_column: Optional[str], group_by: Optional[str],
                        agg_func: str, series_columns: Optional[List[str]], 
                        limit: int) -> Dict[str, Any]:
    """准备图表数据"""
    
    if chart_type == "pie":
        # 饼图：按分组列聚合
        if group_by and group_by in df.columns:
            if y_column and y_column in df.columns:
                grouped = df.groupby(group_by)[y_column].agg(agg_func).reset_index()
                grouped.columns = ["name", "value"]
            else:
                grouped = df[group_by].value_counts().reset_index()
                grouped.columns = ["name", "value"]
            
            grouped = grouped.head(limit)
            data = [{"name": str(row["name"]), "value": float(row["value"])} 
                    for _, row in grouped.iterrows()]
            return {"data": data, "data_count": len(data)}
        else:
            return {"error": "饼图需要指定 group_by 分组列"}
    
    elif chart_type == "scatter":
        # 散点图：需要两个数值列
        if not x_column or not y_column:
            return {"error": "散点图需要指定 x_column 和 y_column"}
        if x_column not in df.columns or y_column not in df.columns:
            return {"error": f"列不存在: {x_column} 或 {y_column}"}
        
        scatter_df = df[[x_column, y_column]].dropna().head(limit * 5)  # 散点图可以多一些点
        data = scatter_df.values.tolist()
        return {
            "data": data, 
            "x_name": x_column, 
            "y_name": y_column,
            "data_count": len(data)
        }
    
    elif chart_type == "radar":
        # 雷达图：多个指标对比
        if not series_columns or len(series_columns) < 3:
            return {"error": "雷达图需要至少3个 series_columns 指标列"}
        
        valid_cols = [c for c in series_columns if c in df.columns and pd.api.types.is_numeric_dtype(df[c])]
        if len(valid_cols) < 3:
            return {"error": "雷达图需要至少3个有效的数值列"}
        
        # 计算每个指标的聚合值
        if group_by and group_by in df.columns:
            # 按分组生成多个雷达系列
            grouped = df.groupby(group_by)[valid_cols].agg(agg_func).head(limit)
            indicators = [{"name": col, "max": float(df[col].max() * 1.2)} for col in valid_cols]
            series_data = []
            for name, row in grouped.iterrows():
                series_data.append({
                    "name": str(name),
                    "value": [float(row[col]) for col in valid_cols]
                })
            return {"indicators": indicators, "series": series_data, "data_count": len(series_data)}
        else:
            # 单系列雷达图
            indicators = [{"name": col, "max": float(df[col].max() * 1.2)} for col in valid_cols]
            values = [float(df[col].agg(agg_func)) for col in valid_cols]
            return {"indicators": indicators, "series": [{"name": "数据", "value": values}], "data_count": 1}
    
    elif chart_type == "funnel":
        # 漏斗图：类似饼图，按值降序排列
        if group_by and group_by in df.columns:
            if y_column and y_column in df.columns:
                grouped = df.groupby(group_by)[y_column].agg(agg_func).reset_index()
                grouped.columns = ["name", "value"]
            else:
                grouped = df[group_by].value_counts().reset_index()
                grouped.columns = ["name", "value"]
            
            grouped = grouped.sort_values("value", ascending=False).head(limit)
            data = [{"name": str(row["name"]), "value": float(row["value"])} 
                    for _, row in grouped.iterrows()]
            return {"data": data, "data_count": len(data)}
        else:
            return {"error": "漏斗图需要指定 group_by 分组列"}
    
    else:
        # bar / line：分类 + 数值
        if not x_column:
            return {"error": f"{chart_type}图需要指定 x_column"}
        if x_column not in df.columns:
            return {"error": f"列 '{x_column}' 不存在"}
        
        # 多系列处理
        if series_columns:
            valid_series = [c for c in series_columns if c in df.columns]
            if not valid_series:
                return {"error": "series_columns 中没有有效的列"}
            
            # 按 x_column 分组，计算每个系列的聚合值
            grouped = df.groupby(x_column)[valid_series].agg(agg_func).head(limit)
            categories = [str(idx) for idx in grouped.index]
            series = [
                {"name": col, "data": grouped[col].tolist()}
                for col in valid_series
            ]
            return {"categories": categories, "series": series, "data_count": len(categories)}
        
        # 单系列处理
        if y_column and y_column in df.columns:
            grouped = df.groupby(x_column)[y_column].agg(agg_func).reset_index()
            grouped.columns = ["category", "value"]
            grouped = grouped.sort_values("value", ascending=False).head(limit)
            categories = [str(c) for c in grouped["category"]]
            values = grouped["value"].tolist()
        else:
            # 仅计数
            grouped = df[x_column].value_counts().head(limit)
            categories = [str(idx) for idx in grouped.index]
            values = grouped.values.tolist()
        
        return {"categories": categories, "values": values, "data_count": len(categories)}


def _build_echart_config(chart_type: str, data: Dict[str, Any], title: str) -> Dict[str, Any]:
    """构建 ECharts 配置"""
    
    # 通用配置
    base_config = {
        "title": {
            "text": title,
            "left": "center",
            "textStyle": {"color": "#e5e7eb"}
        },
        "tooltip": {"trigger": "item" if chart_type in ["pie", "scatter", "funnel"] else "axis"},
        "backgroundColor": "transparent"
    }
    
    if chart_type == "pie":
        return {
            **base_config,
            "legend": {
                "orient": "vertical",
                "left": "left",
                "textStyle": {"color": "#9ca3af"}
            },
            "series": [{
                "type": "pie",
                "radius": ["40%", "70%"],
                "avoidLabelOverlap": True,
                "itemStyle": {
                    "borderRadius": 10,
                    "borderColor": "#1f2937",
                    "borderWidth": 2
                },
                "label": {"color": "#e5e7eb"},
                "emphasis": {
                    "label": {"show": True, "fontSize": 16, "fontWeight": "bold"}
                },
                "data": data["data"]
            }]
        }
    
    elif chart_type == "scatter":
        return {
            **base_config,
            "xAxis": {
                "type": "value",
                "name": data.get("x_name", ""),
                "axisLabel": {"color": "#9ca3af"},
                "axisLine": {"lineStyle": {"color": "#4b5563"}}
            },
            "yAxis": {
                "type": "value",
                "name": data.get("y_name", ""),
                "axisLabel": {"color": "#9ca3af"},
                "axisLine": {"lineStyle": {"color": "#4b5563"}}
            },
            "series": [{
                "type": "scatter",
                "symbolSize": 10,
                "data": data["data"],
                "itemStyle": {"color": "#6366f1"}
            }]
        }
    
    elif chart_type == "radar":
        return {
            **base_config,
            "legend": {
                "data": [s["name"] for s in data["series"]],
                "bottom": 0,
                "textStyle": {"color": "#9ca3af"}
            },
            "radar": {
                "indicator": data["indicators"],
                "axisName": {"color": "#9ca3af"},
                "splitLine": {"lineStyle": {"color": "#4b5563"}},
                "splitArea": {"areaStyle": {"color": ["rgba(99,102,241,0.1)", "rgba(99,102,241,0.05)"]}}
            },
            "series": [{
                "type": "radar",
                "data": data["series"]
            }]
        }
    
    elif chart_type == "funnel":
        return {
            **base_config,
            "legend": {
                "data": [d["name"] for d in data["data"]],
                "bottom": 0,
                "textStyle": {"color": "#9ca3af"}
            },
            "series": [{
                "type": "funnel",
                "left": "10%",
                "width": "80%",
                "label": {"show": True, "position": "inside", "color": "#fff"},
                "labelLine": {"show": False},
                "itemStyle": {"borderColor": "#1f2937", "borderWidth": 1},
                "emphasis": {"label": {"fontSize": 16}},
                "data": data["data"]
            }]
        }
    
    else:
        # bar / line
        config = {
            **base_config,
            "grid": {"left": "3%", "right": "4%", "bottom": "3%", "containLabel": True},
            "xAxis": {
                "type": "category",
                "data": data["categories"],
                "axisLabel": {"color": "#9ca3af", "rotate": 30 if len(data["categories"]) > 8 else 0},
                "axisLine": {"lineStyle": {"color": "#4b5563"}}
            },
            "yAxis": {
                "type": "value",
                "axisLabel": {"color": "#9ca3af"},
                "axisLine": {"lineStyle": {"color": "#4b5563"}},
                "splitLine": {"lineStyle": {"color": "#374151"}}
            }
        }
        
        # 处理多系列
        if "series" in data:
            config["legend"] = {
                "data": [s["name"] for s in data["series"]],
                "bottom": 0,
                "textStyle": {"color": "#9ca3af"}
            }
            config["series"] = [
                {
                    "name": s["name"],
                    "type": chart_type,
                    "data": s["data"],
                    "smooth": chart_type == "line"
                }
                for s in data["series"]
            ]
        else:
            config["series"] = [{
                "type": chart_type,
                "data": data["values"],
                "smooth": chart_type == "line",
                "itemStyle": {"color": "#6366f1"},
                "areaStyle": {"color": "rgba(99,102,241,0.2)"} if chart_type == "line" else None
            }]
        
        return config


# ==================== 写入工具（v2.0 新增） ====================

@tool
def write_cell(
    cell: str,
    value: Any,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """写入单个单元格的值。

    Args:
        cell: 单元格地址，如 "A1", "B2", "C10"
        value: 要写入的值（支持字符串、数字、日期等）
        sheet: 目标工作表名称，默认当前工作表

    Returns:
        操作结果，包含写入前后的值
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
        操作结果，包含写入的行数和单元格数
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
def write_formula(
    cell: str,
    formula: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """写入 Excel 公式到指定单元格。

    注意：公式不会立即计算，需要用 Excel 打开文件后才会计算结果。

    Args:
        cell: 目标单元格地址，如 "C1"
        formula: Excel 公式，如 "SUM(A1:B1)" 或 "=SUM(A1:B1)"（会自动添加 =）
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
            # 读取单元格值
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
        result = loader.insert_rows(row, count, sheet=sheet)
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
def save_file(
    path: Optional[str] = None
) -> Dict[str, Any]:
    """保存当前 Excel 文件到副本。

    注意：修改会保存到工作副本，不会影响原始文件。
    如需保存到原始文件，请使用 save_to_original 工具。

    Args:
        path: 保存路径，默认保存到副本文件。指定新路径则另存为。

    Returns:
        操作结果，包含保存的文件路径
    """
    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}

    if not loader.is_dual_engine():
        return {"error": "当前表不支持保存操作（需要双引擎模式）"}

    try:
        result = loader.save_table(file_path=path)
        save_path = result.get("file_path")
        if path:
            message = f"文件已另存为: {save_path}"
        else:
            message = f"文件已保存到副本: {save_path}"
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
    """将修改保存回原始文件。

    警告：此操作会覆盖原始文件！请确保已备份重要数据。

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
        操作结果，包含实际保存的文件路径
    """
    from pathlib import Path
    import os

    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}

    if not loader.is_dual_engine():
        return {"error": "当前表不支持导出操作（需要双引擎模式）"}

    try:
        # 转换为 Path 对象并获取绝对路径
        export_path_obj = Path(export_path).expanduser().absolute()

        # 确保目标目录存在
        parent_dir = export_path_obj.parent
        if not parent_dir.exists():
            try:
                parent_dir.mkdir(parents=True, exist_ok=True)
            except Exception as dir_error:
                return {"error": f"无法创建目录 {parent_dir}: {str(dir_error)}"}

        # 检查目录是否可写
        if not os.access(str(parent_dir), os.W_OK):
            return {"error": f"没有权限写入目录: {parent_dir}"}

        # 确保文件扩展名正确
        if export_path_obj.suffix.lower() not in ['.xlsx', '.xlsm']:
            export_path_obj = export_path_obj.with_suffix('.xlsx')

        # 执行导出
        result = loader.export_to(str(export_path_obj))
        actual_path = result.get("export_path", str(export_path_obj))

        # 验证文件是否真的创建了
        if not Path(actual_path).exists():
            return {"error": f"导出似乎成功但文件未创建: {actual_path}"}

        return {
            "success": True,
            "export_path": actual_path,
            "message": f"文件已成功导出到: {actual_path}",
            "file_exists": True
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": f"导出失败: {str(e)}"}


@tool
def quick_export(filename_suffix: str = "modified") -> Dict[str, Any]:
    """快速导出文件到用户下载目录或原文件所在目录。

    这是最简单的导出方式，会在用户下载目录创建一个带后缀的新文件。
    如果原文件不是临时文件，则导出到原文件所在目录。
    例如：原文件为 data.xlsx，导出后为 data_modified.xlsx

    Args:
        filename_suffix: 文件名后缀，默认为 "modified"。
                        例如设置为 "final" 则输出 data_final.xlsx

    Returns:
        操作结果，包含导出的文件路径
    """
    from pathlib import Path
    from datetime import datetime
    import os

    loader = get_loader()

    if not loader.is_loaded:
        return {"error": "未加载 Excel 文件"}

    if not loader.is_dual_engine():
        return {"error": "当前表不支持导出操作（需要双引擎模式）"}

    try:
        # 获取原始文件信息
        table_info = loader.get_active_table_info()
        if not table_info:
            return {"error": "无法获取当前表信息"}

        original_path = Path(table_info.original_path)
        original_filename = table_info.filename  # 原始文件名（用户上传时的名字）
        stem = Path(original_filename).stem
        suffix = Path(original_filename).suffix or '.xlsx'

        # 判断是否是临时文件
        is_temp_file = 'temp' in str(original_path).lower() or 'tmp' in str(original_path).lower()

        if is_temp_file:
            # 临时文件：导出到用户下载目录
            # 尝试获取用户下载目录
            if os.name == 'nt':  # Windows
                downloads_dir = Path.home() / "Downloads"
            else:  # Linux/Mac
                downloads_dir = Path.home() / "Downloads"

            if not downloads_dir.exists():
                # 如果下载目录不存在，使用用户主目录
                downloads_dir = Path.home()

            parent_dir = downloads_dir
        else:
            # 非临时文件：导出到原文件所在目录
            parent_dir = original_path.parent

        # 生成新文件名：原文件名_后缀.xlsx
        new_filename = f"{stem}_{filename_suffix}{suffix}"
        export_path = parent_dir / new_filename

        # 如果文件已存在，添加时间戳
        if export_path.exists():
            timestamp = datetime.now().strftime("%H%M%S")
            new_filename = f"{stem}_{filename_suffix}_{timestamp}{suffix}"
            export_path = parent_dir / new_filename

        # 执行导出
        result = loader.export_to(str(export_path))
        actual_path = result.get("export_path", str(export_path))

        # 验证文件是否创建
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
        import traceback
        traceback.print_exc()
        return {"error": f"导出失败: {str(e)}"}


@tool
def get_change_log() -> Dict[str, Any]:
    """获取当前文件的变更记录。

    Returns:
        变更记录列表，包含所有修改操作的详细信息
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


# 写入工具列表
WRITE_TOOLS = [
    write_cell,
    write_range,
    write_formula,
    read_formula,
    insert_rows,
    delete_rows,
    save_file,
    save_to_original,
    export_file,
    get_change_log,
]


# ==================== 格式化工具（v2.0 新增） ====================

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
        color: 字体颜色 (十六进制，如 "FF0000" 红色, "0000FF" 蓝色)
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
            cell_range=cell_range,
            name=name,
            size=size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            sheet=sheet
        )

        if not result.get("success"):
            return {"error": result.get("error", "设置字体失败")}

        settings = []
        if name:
            settings.append(f"字体: {name}")
        if size:
            settings.append(f"字号: {size}")
        if bold:
            settings.append("加粗")
        if italic:
            settings.append("斜体")
        if underline:
            settings.append(f"下划线: {underline}")
        if color:
            settings.append(f"颜色: #{color}")

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
        color: 背景颜色 (十六进制，如 "FFFF00" 黄色, "00FF00" 绿色, "FF0000" 红色)
        fill_type: 填充类型，默认 "solid" (实色填充)
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
        result = doc.set_fill(
            cell_range=cell_range,
            color=color,
            fill_type=fill_type,
            sheet=sheet
        )

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
            cell_range=cell_range,
            horizontal=horizontal,
            vertical=vertical,
            wrap_text=wrap_text,
            sheet=sheet
        )

        if not result.get("success"):
            return {"error": result.get("error", "设置对齐失败")}

        settings = []
        if horizontal:
            settings.append(f"水平: {horizontal}")
        if vertical:
            settings.append(f"垂直: {vertical}")
        if wrap_text:
            settings.append("自动换行")

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
        style: 边框样式 ("thin" 细线, "medium" 中线, "thick" 粗线, "double" 双线, "dashed" 虚线)
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
        result = doc.set_border(
            cell_range=cell_range,
            style=style,
            color=color,
            sheet=sheet
        )

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
    - "#,##0" - 千分位整数 (如 1,234)
    - "#,##0.00" - 千分位两位小数 (如 1,234.56)
    - "0.00%" - 百分比 (如 12.34%)
    - "yyyy-mm-dd" - 日期 (如 2024-01-15)
    - "yyyy/mm/dd" - 日期 (如 2024/01/15)
    - "¥#,##0.00" - 人民币 (如 ¥1,234.56)
    - "$#,##0.00" - 美元 (如 $1,234.56)

    Args:
        cell_range: 单元格或范围，如 "A1" 或 "A1:C10"
        format_code: 数字格式代码
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
        result = doc.set_number_format(
            cell_range=cell_range,
            format_code=format_code,
            sheet=sheet
        )

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
    """一次性设置多种单元格样式（更高效）。

    Args:
        cell_range: 单元格或范围，如 "A1" 或 "A1:C10"
        font_bold: 是否加粗
        font_size: 字号
        font_color: 字体颜色 (十六进制)
        bg_color: 背景颜色 (十六进制)
        horizontal: 水平对齐 ("left", "center", "right")
        border_style: 边框样式 ("thin", "medium", "thick")
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
            cell_range=cell_range,
            font_bold=font_bold,
            font_size=font_size,
            font_color=font_color,
            bg_color=bg_color,
            horizontal=horizontal,
            border_style=border_style,
            number_format=number_format,
            sheet=sheet
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
def merge_cells(
    cell_range: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """合并单元格。

    Args:
        cell_range: 要合并的范围，如 "A1:C3"
        sheet: 工作表名称，默认当前工作表

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

        return {
            "success": True,
            "cell_range": cell_range,
            "message": f"已合并单元格: {cell_range}"
        }
    except Exception as e:
        return {"error": f"合并单元格失败: {str(e)}"}


@tool
def unmerge_cells(
    cell_range: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """取消合并单元格。

    Args:
        cell_range: 要取消合并的范围，如 "A1:C3"
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
        result = doc.unmerge_cells(cell_range=cell_range, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "取消合并失败")}

        return {
            "success": True,
            "cell_range": cell_range,
            "message": f"已取消合并: {cell_range}"
        }
    except Exception as e:
        return {"error": f"取消合并失败: {str(e)}"}


@tool
def set_column_width(
    column: str,
    width: float,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置列宽。

    Args:
        column: 列标识，如 "A", "B", "C"
        width: 列宽（字符数），如 15, 20, 30
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
        result = doc.set_column_width(column=column, width=width, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置列宽失败")}

        return {
            "success": True,
            "column": column.upper(),
            "width": width,
            "message": f"已设置列 {column.upper()} 的宽度为 {width}"
        }
    except Exception as e:
        return {"error": f"设置列宽失败: {str(e)}"}


@tool
def set_row_height(
    row: int,
    height: float,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """设置行高。

    Args:
        row: 行号（从 1 开始）
        height: 行高（磅），如 20, 30, 40
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
        result = doc.set_row_height(row=row, height=height, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "设置行高失败")}

        return {
            "success": True,
            "row": row,
            "height": height,
            "message": f"已设置第 {row} 行的高度为 {height} 磅"
        }
    except Exception as e:
        return {"error": f"设置行高失败: {str(e)}"}


@tool
def auto_fit_column(
    column: str,
    sheet: Optional[str] = None
) -> Dict[str, Any]:
    """自动调整列宽以适应内容。

    Args:
        column: 列标识，如 "A", "B", "C"
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
        result = doc.auto_fit_column(column=column, sheet=sheet)

        if not result.get("success"):
            return {"error": result.get("error", "自动调整列宽失败")}

        return {
            "success": True,
            "column": column.upper(),
            "width": result.get("width"),
            "message": f"已自动调整列 {column.upper()} 的宽度为 {result.get('width')}"
        }
    except Exception as e:
        return {"error": f"自动调整列宽失败: {str(e)}"}


# 格式化工具列表
FORMAT_TOOLS = [
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


# 导出工具列表
ALL_TOOLS = [
    # 查询工具
    filter_data,
    aggregate_data,
    group_and_aggregate,
    search_data,
    get_column_stats,
    get_unique_values,
    get_data_preview,
    # 系统工具
    switch_sheet,
    get_current_time,
    calculate,
    # 可视化
    generate_chart,
    # 写入工具 (v2.0)
    write_cell,
    write_range,
    write_formula,
    read_formula,
    insert_rows,
    delete_rows,
    save_file,
    save_to_original,
    export_file,
    quick_export,
    get_change_log,
    # 格式化工具 (v2.0)
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
