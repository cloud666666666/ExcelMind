"""内置技能定义

将现有工具按技能分组，实现 v2.0 的 Skills 架构。
"""

from typing import List

from .skill_manager import SkillDefinition, SkillCategory, SkillManager
from . import tools


def create_core_query_skill() -> SkillDefinition:
    """创建核心查询技能（始终激活）"""
    return SkillDefinition(
        name="core_query",
        display_name="数据查询",
        description="Excel 数据的基础查询、筛选、搜索、预览和统计功能",
        category=SkillCategory.CORE,
        tools=[
            tools.filter_data,
            tools.search_data,
            tools.get_data_preview,
            tools.get_column_stats,
            tools.get_unique_values,
        ],
        keywords=[
            "查询", "筛选", "过滤", "搜索", "查找", "找出", "找到",
            "显示", "列出", "看看", "有哪些", "多少",
            "预览", "数据", "记录", "行",
            "query", "filter", "search", "find", "show", "list",
        ],
        examples=[
            "帮我筛选出销售额大于1000的记录",
            "查找所有包含'北京'的数据",
            "显示前20行数据",
            "这个表有多少行",
            "A列有哪些唯一值",
        ],
        priority=100,  # 最高优先级
        system_prompt="你可以使用数据查询工具来筛选、搜索和预览 Excel 数据。",
    )


def create_aggregation_skill() -> SkillDefinition:
    """创建聚合分析技能"""
    return SkillDefinition(
        name="aggregation",
        display_name="聚合分析",
        description="数据聚合、分组统计、排序等高级分析功能",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.aggregate_data,
            tools.group_and_aggregate,
            tools.sort_data,
        ],
        keywords=[
            "求和", "平均", "最大", "最小", "总计", "合计", "统计",
            "分组", "汇总", "聚合", "排序", "排名",
            "sum", "avg", "average", "max", "min", "total", "count",
            "group", "aggregate", "sort", "rank",
        ],
        patterns=[
            r"(求和|平均|最大|最小|总计|合计)",
            r"按.+分组",
            r"(升序|降序)排",
        ],
        examples=[
            "计算销售额的总和",
            "按地区分组统计销售额",
            "求出平均价格",
            "按金额降序排列",
        ],
        priority=80,
        requires=["core_query"],  # 依赖核心查询
        system_prompt="你可以使用聚合工具进行求和、平均、分组等统计分析。",
    )


def create_visualization_skill() -> SkillDefinition:
    """创建可视化技能"""
    return SkillDefinition(
        name="visualization",
        display_name="数据可视化",
        description="生成各类图表，包括柱状图、折线图、饼图等",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.generate_chart,
        ],
        keywords=[
            "图表", "图", "柱状图", "折线图", "饼图", "散点图",
            "可视化", "画图", "绘制", "展示",
            "chart", "plot", "graph", "bar", "line", "pie",
        ],
        patterns=[
            r"(画|绘制|生成|创建).*(图|chart)",
            r"(柱状|折线|饼|散点|雷达)图",
        ],
        examples=[
            "画一个销售额的柱状图",
            "生成按月份的折线图",
            "用饼图展示各地区占比",
        ],
        priority=70,
        requires=["core_query"],
        system_prompt="你可以使用图表工具生成各类可视化图表。支持柱状图、折线图、饼图、散点图、雷达图、漏斗图。",
    )


def create_calculation_skill() -> SkillDefinition:
    """创建计算技能"""
    return SkillDefinition(
        name="calculation",
        display_name="数学计算",
        description="执行数学计算和表达式求值",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.calculate,
        ],
        keywords=[
            "计算", "算", "加", "减", "乘", "除", "等于",
            "表达式", "公式", "数学",
            "calculate", "compute", "math",
        ],
        patterns=[
            r"\d+\s*[\+\-\*\/]\s*\d+",
            r"计算.+",
        ],
        examples=[
            "计算 100 * 1.5 + 200",
            "1000 / 4 等于多少",
        ],
        priority=50,
        system_prompt="你可以使用计算工具执行数学运算。",
    )


def create_sheet_management_skill() -> SkillDefinition:
    """创建工作表管理技能"""
    return SkillDefinition(
        name="sheet_management",
        display_name="工作表管理",
        description="切换、管理 Excel 工作表",
        category=SkillCategory.SYSTEM,
        tools=[
            tools.switch_sheet,
        ],
        keywords=[
            "工作表", "sheet", "切换", "表格",
        ],
        patterns=[
            r"切换到.+(表|sheet)",
            r"打开.+(表|sheet)",
        ],
        examples=[
            "切换到 Sheet2",
            "打开销售数据表",
        ],
        priority=60,
        system_prompt="你可以使用工作表管理工具切换不同的工作表。",
    )


def create_utility_skill() -> SkillDefinition:
    """创建实用工具技能"""
    return SkillDefinition(
        name="utility",
        display_name="实用工具",
        description="获取当前时间等实用功能",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.get_current_time,
        ],
        keywords=[
            "时间", "日期", "今天", "现在",
            "time", "date", "today", "now",
        ],
        examples=[
            "现在几点了",
            "今天是几号",
        ],
        priority=30,
    )


def create_modification_skill() -> SkillDefinition:
    """创建数据修改技能（v2.0 新增）

    包含写入单元格、批量写入、行列操作、保存文件等功能。
    """
    return SkillDefinition(
        name="modification",
        display_name="数据修改",
        description="写入、修改、删除 Excel 数据，包括单元格写入、批量写入、行列操作、保存文件",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.write_cell,
            tools.write_range,
            tools.insert_rows,
            tools.delete_rows,
            tools.save_file,
            tools.save_to_original,
            tools.export_file,
            tools.quick_export,
            tools.get_change_log,
        ],
        keywords=[
            "写入", "修改", "更新", "删除", "添加", "插入", "加上",
            "保存", "另存", "导出", "覆盖", "原始文件",
            "末尾", "结尾", "最后", "新增", "追加",
            "write", "update", "delete", "insert", "save", "export", "append",
        ],
        patterns=[
            r"(写入|修改|更新|删除).+",
            r"把.+(改成|设为|设置为)",
            r"在.+(添加|插入|加上)",
            r"(末尾|结尾|最后).*(加上|添加|写入)",
            r"(加上|添加).*(合计|总计|汇总)",
            r"保存(文件|表格|到原始)?",
            r"导出(到|为)?",
        ],
        examples=[
            "把 A1 单元格写入 100",
            "在 A1 开始写入数据",
            "删除第 5 行",
            "插入 3 行",
            "保存文件",
            "保存到原始文件",
            "导出到新文件",
        ],
        priority=75,
        requires=["core_query"],
        system_prompt="""你可以使用数据修改工具来写入和修改 Excel 数据。
注意：
- 所有修改默认保存到工作副本，不会影响原始文件
- 使用 save_file 保存到副本
- 使用 save_to_original 覆盖原始文件（慎用）
- 使用 quick_export 快速导出到原文件所在目录（推荐）
- 使用 export_file 导出到指定位置
- 可以使用 get_change_log 查看所有修改记录""",
    )


def create_formula_skill() -> SkillDefinition:
    """创建公式技能（v2.0 新增）

    包含读取和写入 Excel 公式的功能。
    """
    return SkillDefinition(
        name="formula",
        display_name="Excel 公式",
        description="读取和写入 Excel 公式，支持各种 Excel 函数",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.write_formula,
            tools.read_formula,
        ],
        keywords=[
            "公式", "函数", "SUM", "AVERAGE", "COUNT", "MAX", "MIN",
            "VLOOKUP", "IF", "SUMIF", "COUNTIF",
            "formula", "function",
        ],
        patterns=[
            r"=\w+\(",
            r"(添加|写入|设置).*(公式|函数)",
            r"(读取|查看|获取).*(公式|函数)",
        ],
        examples=[
            "在 C1 添加求和公式 =SUM(A1:B1)",
            "读取 D1 单元格的公式",
            "这个单元格用的是什么公式",
            "写入 AVERAGE 函数",
        ],
        priority=70,
        requires=["core_query"],
        system_prompt="""你可以读取和写入 Excel 公式。
注意：
- 公式将在 Excel 中打开时计算，不会立即显示结果
- 支持所有标准 Excel 函数（SUM, AVERAGE, COUNT, IF, VLOOKUP 等）
- 写入公式时可以省略开头的 = 号，系统会自动添加""",
    )


def create_formatting_skill() -> SkillDefinition:
    """创建格式化技能（v2.0 新增）"""
    return SkillDefinition(
        name="formatting",
        display_name="格式设置",
        description="设置单元格格式、样式、字体、颜色、边框、合并单元格、调整行高列宽等",
        category=SkillCategory.ON_DEMAND,
        tools=[
            tools.set_font,
            tools.set_fill,
            tools.set_alignment,
            tools.set_border,
            tools.set_number_format,
            tools.set_cell_style,
            tools.merge_cells,
            tools.unmerge_cells,
            tools.set_column_width,
            tools.set_row_height,
            tools.auto_fit_column,
        ],
        keywords=[
            "格式", "样式", "字体", "颜色", "边框", "对齐",
            "加粗", "斜体", "下划线", "背景色", "填充",
            "合并", "取消合并", "列宽", "行高", "自动调整",
            "居中", "居左", "居右",
            "format", "style", "font", "color", "border", "merge",
        ],
        patterns=[
            r"(设置|修改).*(格式|样式|字体|颜色|背景|边框)",
            r"把.+(加粗|变色|居中|合并)",
            r"(合并|取消合并).*(单元格|格子)",
            r"(调整|设置).*(列宽|行高)",
        ],
        examples=[
            "把标题行加粗",
            "设置 A 列为红色",
            "给表格添加边框",
            "把 A1:C1 居中",
            "合并 A1:C1 单元格",
            "设置第一行背景为黄色",
            "调整 A 列宽度",
        ],
        priority=65,
        requires=["core_query"],
        system_prompt="""你可以设置单元格的格式和样式，包括：
- 字体设置：字体名称、大小、加粗、斜体、颜色
- 填充设置：背景颜色
- 对齐设置：水平对齐、垂直对齐、自动换行
- 边框设置：边框样式、颜色
- 数字格式：千分位、百分比、日期、货币等
- 合并/取消合并单元格
- 调整列宽、行高
注意：格式化操作需要保存文件后才能在 Excel 中看到效果。""",
    )


# ==================== 技能注册 ====================

def register_builtin_skills(manager: SkillManager) -> None:
    """注册所有内置技能

    Args:
        manager: SkillManager 实例
    """
    # 核心技能
    manager.register(create_core_query_skill())

    # 分析技能
    manager.register(create_aggregation_skill())
    manager.register(create_visualization_skill())
    manager.register(create_calculation_skill())

    # 系统技能
    manager.register(create_sheet_management_skill())
    manager.register(create_utility_skill())

    # v2.0 新增技能（部分工具待实现）
    manager.register(create_modification_skill())
    manager.register(create_formula_skill())
    manager.register(create_formatting_skill())


def create_skill_manager() -> SkillManager:
    """创建并初始化 SkillManager

    Returns:
        已注册内置技能的 SkillManager 实例
    """
    manager = SkillManager()
    register_builtin_skills(manager)
    return manager


# ==================== 全局实例 ====================

_skill_manager: SkillManager = None


def get_skill_manager() -> SkillManager:
    """获取全局 SkillManager 实例"""
    global _skill_manager
    if _skill_manager is None:
        _skill_manager = create_skill_manager()
    return _skill_manager


def reset_skill_manager() -> None:
    """重置全局 SkillManager 实例"""
    global _skill_manager
    _skill_manager = create_skill_manager()
