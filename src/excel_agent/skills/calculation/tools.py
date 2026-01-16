"""数学计算工具"""

import math
from typing import Any, Dict, List
from langchain_core.tools import tool


@tool
def calculate(expressions: List[str]) -> Dict[str, Any]:
    """执行数学计算。

    Args:
        expressions: 数学表达式列表，例如 ["(100+200)*0.5", "500/2"]

    Returns:
        每个表达式的计算结果
    """
    results = {}

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
            if any(char in expr for char in ["__", "import", "eval", "exec", "open"]):
                results[expr] = "Error: Unsafe expression"
                continue

            result = eval(expr, {"__builtins__": None}, safe_env)
            results[expr] = result
        except Exception as e:
            results[expr] = f"Error: {str(e)}"

    return {"results": results}


# 导出工具列表
TOOLS = [calculate]
