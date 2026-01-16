"""数据可视化工具"""

from typing import Any, Dict, List, Optional
import pandas as pd
from langchain_core.tools import tool

from .._common import _get_filter_mask
from ...excel_loader import get_loader


@tool
def generate_chart(
    chart_type: Optional[str] = None,
    x_column: Optional[str] = None,
    y_column: Optional[str] = None,
    agg_column: Optional[str] = None,
    group_by: Optional[str] = None,
    agg_func: str = "sum",
    title: str = "",
    filters: Optional[List[Dict[str, Any]]] = None,
    series_columns: Optional[List[str]] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """生成 ECharts 可视化图表配置。

    Args:
        chart_type: 图表类型: bar, line, pie, scatter, radar, funnel。为空时自动推荐。
        x_column: X轴数据列名
        y_column: Y轴数据列名
        agg_column: 聚合列名（y_column 的别名）
        group_by: 分组列名（用于饼图和多系列图）
        agg_func: 聚合函数: sum, mean, count, min, max
        title: 图表标题
        filters: 筛选条件列表
        series_columns: 多系列Y轴列名列表
        limit: 数据点数量限制，默认20

    Returns:
        包含 ECharts 配置的字典
    """
    if agg_column and not y_column:
        y_column = agg_column

    loader = get_loader()
    df = loader.dataframe.copy()

    # 应用筛选条件
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

    if len(df) == 0:
        return {"error": "筛选后无数据，无法生成图表"}

    # 自动推荐图表类型
    def recommend_chart_type() -> str:
        if group_by and y_column:
            unique_groups = df[group_by].nunique() if group_by in df.columns else 0
            if unique_groups <= 8:
                return "pie"
            return "bar"

        if x_column and y_column:
            x_dtype = df[x_column].dtype if x_column in df.columns else None
            y_dtype = df[y_column].dtype if y_column in df.columns else None

            if pd.api.types.is_numeric_dtype(x_dtype) and pd.api.types.is_numeric_dtype(y_dtype):
                return "scatter"
            if pd.api.types.is_datetime64_any_dtype(x_dtype):
                return "line"
            return "bar"

        if group_by:
            return "pie"
        return "bar"

    final_chart_type = chart_type if chart_type and chart_type != "auto" else recommend_chart_type()

    try:
        chart_data = _prepare_chart_data(df, final_chart_type, x_column, y_column,
                                         group_by, agg_func, series_columns, limit)
        if "error" in chart_data:
            return chart_data

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
        if not x_column or not y_column:
            return {"error": "散点图需要指定 x_column 和 y_column"}
        if x_column not in df.columns or y_column not in df.columns:
            return {"error": f"列不存在: {x_column} 或 {y_column}"}

        scatter_df = df[[x_column, y_column]].dropna().head(limit * 5)
        data = scatter_df.values.tolist()
        return {"data": data, "x_name": x_column, "y_name": y_column, "data_count": len(data)}

    elif chart_type == "radar":
        if not series_columns or len(series_columns) < 3:
            return {"error": "雷达图需要至少3个 series_columns 指标列"}

        valid_cols = [c for c in series_columns if c in df.columns and pd.api.types.is_numeric_dtype(df[c])]
        if len(valid_cols) < 3:
            return {"error": "雷达图需要至少3个有效的数值列"}

        if group_by and group_by in df.columns:
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
            indicators = [{"name": col, "max": float(df[col].max() * 1.2)} for col in valid_cols]
            values = [float(df[col].agg(agg_func)) for col in valid_cols]
            return {"indicators": indicators, "series": [{"name": "数据", "value": values}], "data_count": 1}

    elif chart_type == "funnel":
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
        # bar / line
        if not x_column:
            return {"error": f"{chart_type}图需要指定 x_column"}
        if x_column not in df.columns:
            return {"error": f"列 '{x_column}' 不存在"}

        if series_columns:
            valid_series = [c for c in series_columns if c in df.columns]
            if not valid_series:
                return {"error": "series_columns 中没有有效的列"}

            grouped = df.groupby(x_column)[valid_series].agg(agg_func).head(limit)
            categories = [str(idx) for idx in grouped.index]
            series = [{"name": col, "data": grouped[col].tolist()} for col in valid_series]
            return {"categories": categories, "series": series, "data_count": len(categories)}

        if y_column and y_column in df.columns:
            grouped = df.groupby(x_column)[y_column].agg(agg_func).reset_index()
            grouped.columns = ["category", "value"]
            grouped = grouped.sort_values("value", ascending=False).head(limit)
            categories = [str(c) for c in grouped["category"]]
            values = grouped["value"].tolist()
        else:
            grouped = df[x_column].value_counts().head(limit)
            categories = [str(idx) for idx in grouped.index]
            values = grouped.values.tolist()

        return {"categories": categories, "values": values, "data_count": len(categories)}


def _build_echart_config(chart_type: str, data: Dict[str, Any], title: str) -> Dict[str, Any]:
    """构建 ECharts 配置"""

    base_config = {
        "title": {"text": title, "left": "center", "textStyle": {"color": "#e5e7eb"}},
        "tooltip": {"trigger": "item" if chart_type in ["pie", "scatter", "funnel"] else "axis"},
        "backgroundColor": "transparent"
    }

    if chart_type == "pie":
        return {
            **base_config,
            "legend": {"orient": "vertical", "left": "left", "textStyle": {"color": "#9ca3af"}},
            "series": [{
                "type": "pie",
                "radius": ["40%", "70%"],
                "avoidLabelOverlap": True,
                "itemStyle": {"borderRadius": 10, "borderColor": "#1f2937", "borderWidth": 2},
                "label": {"color": "#e5e7eb"},
                "emphasis": {"label": {"show": True, "fontSize": 16, "fontWeight": "bold"}},
                "data": data["data"]
            }]
        }

    elif chart_type == "scatter":
        return {
            **base_config,
            "xAxis": {
                "type": "value", "name": data.get("x_name", ""),
                "axisLabel": {"color": "#9ca3af"}, "axisLine": {"lineStyle": {"color": "#4b5563"}}
            },
            "yAxis": {
                "type": "value", "name": data.get("y_name", ""),
                "axisLabel": {"color": "#9ca3af"}, "axisLine": {"lineStyle": {"color": "#4b5563"}}
            },
            "series": [{"type": "scatter", "symbolSize": 10, "data": data["data"], "itemStyle": {"color": "#6366f1"}}]
        }

    elif chart_type == "radar":
        return {
            **base_config,
            "legend": {"data": [s["name"] for s in data["series"]], "bottom": 0, "textStyle": {"color": "#9ca3af"}},
            "radar": {
                "indicator": data["indicators"],
                "axisName": {"color": "#9ca3af"},
                "splitLine": {"lineStyle": {"color": "#4b5563"}},
                "splitArea": {"areaStyle": {"color": ["rgba(99,102,241,0.1)", "rgba(99,102,241,0.05)"]}}
            },
            "series": [{"type": "radar", "data": data["series"]}]
        }

    elif chart_type == "funnel":
        return {
            **base_config,
            "legend": {"data": [d["name"] for d in data["data"]], "bottom": 0, "textStyle": {"color": "#9ca3af"}},
            "series": [{
                "type": "funnel", "left": "10%", "width": "80%",
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
                "type": "category", "data": data["categories"],
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

        if "series" in data:
            config["legend"] = {"data": [s["name"] for s in data["series"]], "bottom": 0, "textStyle": {"color": "#9ca3af"}}
            config["series"] = [
                {"name": s["name"], "type": chart_type, "data": s["data"], "smooth": chart_type == "line"}
                for s in data["series"]
            ]
        else:
            config["series"] = [{
                "type": chart_type, "data": data["values"], "smooth": chart_type == "line",
                "itemStyle": {"color": "#6366f1"},
                "areaStyle": {"color": "rgba(99,102,241,0.2)"} if chart_type == "line" else None
            }]

        return config


# 导出工具列表
TOOLS = [generate_chart]
