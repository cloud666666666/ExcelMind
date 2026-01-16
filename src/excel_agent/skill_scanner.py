"""SkillScanner - 技能文件扫描器

类似 Claude Code 的技能扫描机制：
1. 启动时扫描 skills/ 目录
2. 只加载元数据（name, description, keywords）到内存
3. 需要时才懒加载完整的 skill 内容
4. 支持从技能目录动态加载 tools.py

Skill 目录格式（Claude Code 风格）:
```
skills/
  skill_name/
    SKILL.md       # 元数据（YAML frontmatter）+ 详细文档
    tools.py       # 技能工具实现（可选）
    LICENSE.txt    # 许可证（可选）
```

SKILL.md 文件格式:
```markdown
---
name: skill_name
display_name: 技能显示名称
description: 简短描述
category: core | on_demand | system
priority: 80
keywords: [关键词1, 关键词2]
patterns: [正则模式]
tools: [tool_name_1, tool_name_2]
requires: [依赖技能]
examples: [示例1, 示例2]
---

# 技能详细文档

详细的使用说明和系统提示内容...
```
"""

import importlib
import importlib.util
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set

import yaml

from .skill_manager import SkillCategory, SkillDefinition


@dataclass
class SkillMetadata:
    """技能元数据（轻量级，用于列表展示）"""
    name: str
    display_name: str
    description: str
    category: SkillCategory
    priority: int = 0
    keywords: List[str] = field(default_factory=list)
    patterns: List[str] = field(default_factory=list)
    tool_names: List[str] = field(default_factory=list)
    file_path: Optional[str] = None  # 文件路径，用于懒加载

    def to_list_item(self) -> str:
        """生成列表展示格式（节省 token）"""
        return f"- {self.name}: {self.description}"


@dataclass
class SkillFileContent:
    """技能文件完整内容"""
    metadata: SkillMetadata
    system_prompt: str = ""
    examples: List[str] = field(default_factory=list)
    requires: List[str] = field(default_factory=list)
    conflicts: List[str] = field(default_factory=list)


class SkillScanner:
    """技能扫描器

    核心设计思想：
    1. 启动时快速扫描，只加载元数据
    2. 生成技能列表供系统提示使用（节省 token）
    3. 需要时才加载完整内容
    4. 支持从技能目录动态加载 tools.py

    使用示例:
    ```python
    scanner = SkillScanner("skills/")
    scanner.scan()  # 扫描所有技能文件

    # 获取技能列表（用于系统提示）
    skill_list = scanner.get_skill_list_prompt()

    # 懒加载完整内容
    full_content = scanner.load_full_skill("aggregation")

    # 加载技能目录的工具
    tools = scanner.load_skill_tools("core_query")
    ```
    """

    def __init__(self, skills_dir: str = None):
        """初始化扫描器

        Args:
            skills_dir: 技能文件目录路径，默认为模块同级 skills/ 目录
        """
        if skills_dir is None:
            # 默认使用模块同级的 skills 目录
            module_dir = Path(__file__).parent
            skills_dir = module_dir / "skills"

        self.skills_dir = Path(skills_dir)
        self._metadata_cache: Dict[str, SkillMetadata] = {}
        self._full_cache: Dict[str, SkillFileContent] = {}
        self._tools_cache: Dict[str, List[Callable]] = {}
        self._scanned = False

    def scan(self, force: bool = False) -> int:
        """扫描技能目录

        Args:
            force: 是否强制重新扫描

        Returns:
            发现的技能数量
        """
        if self._scanned and not force:
            return len(self._metadata_cache)

        self._metadata_cache.clear()
        self._full_cache.clear()

        if not self.skills_dir.exists():
            print(f"[SkillScanner] 技能目录不存在: {self.skills_dir}")
            return 0

        count = 0

        # 1. 扫描目录格式的技能（Claude Code 风格）
        for skill_dir in self.skills_dir.iterdir():
            if skill_dir.is_dir():
                skill_md = skill_dir / "SKILL.md"
                if skill_md.exists():
                    try:
                        metadata = self._load_metadata_from_skill_md(skill_md)
                        if metadata:
                            self._metadata_cache[metadata.name] = metadata
                            count += 1
                    except Exception as e:
                        print(f"[SkillScanner] 加载技能目录失败 {skill_dir}: {e}")

        # 2. 向后兼容：扫描 YAML 文件格式
        for file_path in self.skills_dir.glob("*.yaml"):
            try:
                metadata = self._load_metadata_from_yaml(file_path)
                if metadata and metadata.name not in self._metadata_cache:
                    self._metadata_cache[metadata.name] = metadata
                    count += 1
            except Exception as e:
                print(f"[SkillScanner] 加载技能文件失败 {file_path}: {e}")

        for file_path in self.skills_dir.glob("*.yml"):
            try:
                metadata = self._load_metadata_from_yaml(file_path)
                if metadata and metadata.name not in self._metadata_cache:
                    self._metadata_cache[metadata.name] = metadata
                    count += 1
            except Exception as e:
                print(f"[SkillScanner] 加载技能文件失败 {file_path}: {e}")

        self._scanned = True
        print(f"[SkillScanner] 扫描完成，发现 {count} 个技能")
        return count

    def load_skill_tools(self, skill_name: str) -> List[Callable]:
        """加载技能目录的工具

        从技能目录的 tools.py 文件动态加载工具函数。

        Args:
            skill_name: 技能名称

        Returns:
            工具函数列表
        """
        # 检查缓存
        if skill_name in self._tools_cache:
            return self._tools_cache[skill_name]

        # 查找技能目录
        skill_dir = self.skills_dir / skill_name
        tools_file = skill_dir / "tools.py"

        if not tools_file.exists():
            print(f"[SkillScanner] 技能 {skill_name} 没有 tools.py 文件")
            return []

        try:
            # 动态加载模块
            spec = importlib.util.spec_from_file_location(
                f"excel_agent.skills.{skill_name}.tools",
                tools_file
            )
            if spec is None or spec.loader is None:
                print(f"[SkillScanner] 无法加载技能 {skill_name} 的 tools.py")
                return []

            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # 获取 TOOLS 列表
            if hasattr(module, 'TOOLS'):
                tools = module.TOOLS
                self._tools_cache[skill_name] = tools
                print(f"[SkillScanner] 已加载技能 {skill_name} 的 {len(tools)} 个工具")
                return tools
            else:
                print(f"[SkillScanner] 技能 {skill_name} 的 tools.py 没有 TOOLS 列表")
                return []

        except Exception as e:
            print(f"[SkillScanner] 加载技能 {skill_name} 的工具失败: {e}")
            return []

    def get_skill_tools_registry(self, skill_name: str) -> Dict[str, Callable]:
        """获取技能的工具注册表

        Args:
            skill_name: 技能名称

        Returns:
            工具注册表 {tool_name: tool_function}
        """
        tools = self.load_skill_tools(skill_name)
        registry = {}
        for tool in tools:
            if hasattr(tool, 'name'):
                registry[tool.name] = tool
        return registry

    def get_all_tools_registry(self) -> Dict[str, Callable]:
        """获取所有技能的工具注册表

        Returns:
            合并后的工具注册表 {tool_name: tool_function}
        """
        registry = {}

        for skill_name in self._metadata_cache.keys():
            skill_registry = self.get_skill_tools_registry(skill_name)
            registry.update(skill_registry)

        return registry

    def _parse_yaml_frontmatter(self, content: str) -> tuple[dict, str]:
        """解析 SKILL.md 的 YAML frontmatter 和 markdown 内容

        Args:
            content: SKILL.md 文件内容

        Returns:
            (frontmatter_dict, markdown_body)
        """
        if not content.startswith('---'):
            return {}, content

        # 找到第二个 ---
        end_idx = content.find('---', 3)
        if end_idx == -1:
            return {}, content

        frontmatter_str = content[3:end_idx].strip()
        body = content[end_idx + 3:].strip()

        try:
            frontmatter = yaml.safe_load(frontmatter_str) or {}
        except Exception:
            frontmatter = {}

        return frontmatter, body

    def _load_metadata_from_skill_md(self, skill_md_path: Path) -> Optional[SkillMetadata]:
        """从 SKILL.md 文件加载元数据（Claude Code 风格）"""
        with open(skill_md_path, 'r', encoding='utf-8') as f:
            content = f.read()

        data, _ = self._parse_yaml_frontmatter(content)

        if not data or 'name' not in data:
            # 如果没有 name，使用目录名
            data['name'] = skill_md_path.parent.name

        return self._create_metadata(data, str(skill_md_path))

    def _load_metadata_from_yaml(self, file_path: Path) -> Optional[SkillMetadata]:
        """从 YAML 文件加载元数据（向后兼容）"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)

        if not data or 'name' not in data:
            return None

        return self._create_metadata(data, str(file_path))

    def _create_metadata(self, data: dict, file_path: str) -> SkillMetadata:
        """从解析后的数据创建 SkillMetadata"""
        # 解析 category
        category_str = data.get('category', 'on_demand')
        try:
            category = SkillCategory(category_str)
        except ValueError:
            category = SkillCategory.ON_DEMAND

        return SkillMetadata(
            name=data['name'],
            display_name=data.get('display_name', data['name']),
            description=data.get('description', ''),
            category=category,
            priority=data.get('priority', 0),
            keywords=data.get('keywords', []),
            patterns=data.get('patterns', []),
            tool_names=data.get('tools', []),
            file_path=file_path,
        )

    def load_full_skill(self, skill_name: str) -> Optional[SkillFileContent]:
        """懒加载技能完整内容

        Args:
            skill_name: 技能名称

        Returns:
            技能完整内容，不存在返回 None
        """
        # 检查缓存
        if skill_name in self._full_cache:
            return self._full_cache[skill_name]

        # 检查元数据是否存在
        metadata = self._metadata_cache.get(skill_name)
        if not metadata or not metadata.file_path:
            return None

        # 根据文件类型加载完整内容
        try:
            file_path = Path(metadata.file_path)

            if file_path.name == "SKILL.md":
                # Claude Code 风格：从 SKILL.md 加载
                content = self._load_full_from_skill_md(file_path, metadata)
            else:
                # 向后兼容：从 YAML 文件加载
                content = self._load_full_from_yaml(file_path, metadata)

            if content:
                self._full_cache[skill_name] = content
            return content

        except Exception as e:
            print(f"[SkillScanner] 加载技能完整内容失败 {skill_name}: {e}")
            return None

    def _load_full_from_skill_md(self, file_path: Path, metadata: SkillMetadata) -> SkillFileContent:
        """从 SKILL.md 加载完整内容"""
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        data, markdown_body = self._parse_yaml_frontmatter(content)

        return SkillFileContent(
            metadata=metadata,
            system_prompt=markdown_body,  # markdown 内容作为 system_prompt
            examples=data.get('examples', []),
            requires=data.get('requires', []),
            conflicts=data.get('conflicts', []),
        )

    def _load_full_from_yaml(self, file_path: Path, metadata: SkillMetadata) -> SkillFileContent:
        """从 YAML 文件加载完整内容（向后兼容）"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)

        return SkillFileContent(
            metadata=metadata,
            system_prompt=data.get('system_prompt', ''),
            examples=data.get('examples', []),
            requires=data.get('requires', []),
            conflicts=data.get('conflicts', []),
        )

    def get_skill_list_prompt(self) -> str:
        """生成技能列表提示（用于系统提示，节省 token）

        Returns:
            格式化的技能列表字符串
        """
        if not self._metadata_cache:
            return "暂无可用技能。"

        # 按 category 分组
        by_category: Dict[SkillCategory, List[SkillMetadata]] = {}
        for metadata in self._metadata_cache.values():
            if metadata.category not in by_category:
                by_category[metadata.category] = []
            by_category[metadata.category].append(metadata)

        lines = ["可用技能列表："]

        # 按优先级显示
        for category in [SkillCategory.CORE, SkillCategory.ON_DEMAND, SkillCategory.SYSTEM]:
            skills = by_category.get(category, [])
            if skills:
                # 按优先级排序
                skills.sort(key=lambda x: x.priority, reverse=True)
                for skill in skills:
                    lines.append(skill.to_list_item())

        return "\n".join(lines)

    def get_all_metadata(self) -> List[SkillMetadata]:
        """获取所有技能元数据"""
        return list(self._metadata_cache.values())

    def get_metadata(self, skill_name: str) -> Optional[SkillMetadata]:
        """获取指定技能的元数据"""
        return self._metadata_cache.get(skill_name)

    def to_skill_definition(
        self,
        skill_name: str,
        tools_registry: Dict[str, Callable] = None
    ) -> Optional[SkillDefinition]:
        """将技能文件转换为 SkillDefinition

        Args:
            skill_name: 技能名称
            tools_registry: 工具注册表 {tool_name: tool_function}
                           如果为 None，则尝试从技能目录加载工具

        Returns:
            SkillDefinition 实例
        """
        content = self.load_full_skill(skill_name)
        if not content:
            return None

        metadata = content.metadata

        # 尝试从技能目录加载工具
        skill_tools = self.load_skill_tools(skill_name)

        # 根据 tool_names 获取实际的工具函数
        tools = []

        if skill_tools:
            # 优先使用技能目录的工具
            skill_tools_registry = {t.name: t for t in skill_tools if hasattr(t, 'name')}
            for tool_name in metadata.tool_names:
                if tool_name in skill_tools_registry:
                    tools.append(skill_tools_registry[tool_name])
                elif tools_registry and tool_name in tools_registry:
                    tools.append(tools_registry[tool_name])
                else:
                    print(f"[SkillScanner] 警告: 技能 {skill_name} 引用的工具 {tool_name} 不存在")
        elif tools_registry:
            # 回退到传入的注册表
            for tool_name in metadata.tool_names:
                if tool_name in tools_registry:
                    tools.append(tools_registry[tool_name])
                else:
                    print(f"[SkillScanner] 警告: 技能 {skill_name} 引用的工具 {tool_name} 不存在")
        else:
            print(f"[SkillScanner] 警告: 技能 {skill_name} 没有可用的工具")

        return SkillDefinition(
            name=metadata.name,
            display_name=metadata.display_name,
            description=metadata.description,
            category=metadata.category,
            tools=tools,
            keywords=metadata.keywords,
            patterns=metadata.patterns,
            examples=content.examples,
            priority=metadata.priority,
            requires=content.requires,
            conflicts=content.conflicts,
            system_prompt=content.system_prompt,
        )


# ==================== 全局实例 ====================

_skill_scanner: Optional[SkillScanner] = None


def get_skill_scanner() -> SkillScanner:
    """获取全局 SkillScanner 实例"""
    global _skill_scanner
    if _skill_scanner is None:
        _skill_scanner = SkillScanner()
        _skill_scanner.scan()
    return _skill_scanner


def reset_skill_scanner() -> None:
    """重置全局 SkillScanner 实例"""
    global _skill_scanner
    _skill_scanner = None
