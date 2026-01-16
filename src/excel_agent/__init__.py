"""Excel 智能问数 Agent

v2.0: 支持双引擎模式 (pandas + openpyxl) 和 Skills 动态路由
v2.1: 新增 Claude Code 风格的技能扫描系统
"""

__version__ = "2.1.0"

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
from .skill_scanner import (
    SkillScanner,
    SkillMetadata,
    SkillFileContent,
    get_skill_scanner,
    reset_skill_scanner,
)
from .skill_loader import (
    SkillLoader,
    get_skill_loader,
    reset_skill_loader,
    build_tools_registry,
)
from .tools import (
    get_tools_registry,
    get_tools_by_names,
    ALL_TOOLS,
    SKILL_TOOLS,
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
    # Skills 管理
    "SkillManager",
    "SkillDefinition",
    "SkillCategory",
    "IntentMatch",
    # Skills 扫描（Claude Code 风格）
    "SkillScanner",
    "SkillMetadata",
    "SkillFileContent",
    "get_skill_scanner",
    "reset_skill_scanner",
    "SkillLoader",
    "get_skill_loader",
    "reset_skill_loader",
    "build_tools_registry",
    # 工具
    "get_tools_registry",
    "get_tools_by_names",
    "ALL_TOOLS",
    "SKILL_TOOLS",
]
