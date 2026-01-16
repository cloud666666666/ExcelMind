"""实用工具"""

from datetime import datetime
from typing import Any, Dict
from langchain_core.tools import tool


@tool
def get_current_time() -> Dict[str, Any]:
    """获取当前系统时间。

    Returns:
        当前时间信息
    """
    now = datetime.now()
    return {
        "current_time": now.strftime("%Y-%m-%d %H:%M:%S"),
        "weekday": now.strftime("%A"),
        "timestamp": now.timestamp()
    }


# 导出工具列表
TOOLS = [get_current_time]
