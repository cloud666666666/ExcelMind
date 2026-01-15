"""Excel 智能问数 Agent

v2.0: 支持双引擎模式 (pandas + openpyxl) 和 Skills 动态路由
"""

__version__ = "2.0.0"

from .excel_document import ExcelDocument, Change, ChangeType
from .excel_loader import (
    ExcelLoader,
    MultiExcelLoader,
    TableInfo,
    get_loader,
    reset_loader,
)
from .skill_manager import (
    SkillManager,
    SkillDefinition,
    SkillCategory,
    IntentMatch,
)
from .skills import (
    get_skill_manager,
    reset_skill_manager,
    create_skill_manager,
    register_builtin_skills,
)

__all__ = [
    # 双引擎
    "ExcelDocument",
    "ExcelLoader",
    "MultiExcelLoader",
    "TableInfo",
    "Change",
    "ChangeType",
    "get_loader",
    "reset_loader",
    # Skills
    "SkillManager",
    "SkillDefinition",
    "SkillCategory",
    "IntentMatch",
    "get_skill_manager",
    "reset_skill_manager",
    "create_skill_manager",
    "register_builtin_skills",
]
