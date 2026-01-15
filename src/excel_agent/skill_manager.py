"""SkillManager - 技能动态路由管理器

实现基于意图的动态工具加载，减少上下文 Token 消耗。
"""

import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable, Dict, List, Optional, Set

from langchain_core.tools import BaseTool


class SkillCategory(Enum):
    """技能类别"""
    CORE = "core"           # 核心技能，始终激活
    ON_DEMAND = "on_demand" # 按需加载
    SYSTEM = "system"       # 系统技能


@dataclass
class SkillDefinition:
    """技能定义

    一个 Skill 是一组相关工具的逻辑集合，面向特定任务领域。
    """
    name: str                              # 唯一标识: "data_query"
    display_name: str                      # 显示名称: "数据查询"
    description: str                       # 语义描述，用于意图匹配
    category: SkillCategory                # 技能类别
    tools: List[Callable] = field(default_factory=list)  # 包含的工具函数

    # 激活配置
    keywords: List[str] = field(default_factory=list)     # 强触发关键词
    patterns: List[str] = field(default_factory=list)     # 正则匹配模式
    examples: List[str] = field(default_factory=list)     # 触发示例（用于语义匹配）

    # 优先级与依赖
    priority: int = 0                      # 优先级，数值越高优先级越高
    requires: List[str] = field(default_factory=list)     # 前置依赖的技能
    conflicts: List[str] = field(default_factory=list)    # 互斥的技能

    # 额外系统提示
    system_prompt: str = ""                # 激活时附加的系统提示

    def __hash__(self):
        return hash(self.name)

    def __eq__(self, other):
        if isinstance(other, SkillDefinition):
            return self.name == other.name
        return False


class IntentMatch:
    """意图匹配结果"""

    def __init__(
        self,
        skill: SkillDefinition,
        score: float,
        match_type: str,
        matched_text: str = ""
    ):
        self.skill = skill
        self.score = score          # 匹配分数 0-1
        self.match_type = match_type  # "keyword", "pattern", "semantic"
        self.matched_text = matched_text

    def __repr__(self):
        return f"IntentMatch({self.skill.name}, score={self.score:.2f}, type={self.match_type})"


class SkillManager:
    """技能管理器

    负责技能注册、意图路由、工具动态加载。

    使用示例:
    ```python
    manager = SkillManager()

    # 注册技能
    manager.register(query_skill)
    manager.register(modification_skill)

    # 根据用户输入解析需要的技能
    skills = manager.resolve("帮我筛选出销售额大于1000的记录")

    # 获取激活的工具列表
    tools = manager.get_active_tools()
    ```
    """

    def __init__(self):
        self._registry: Dict[str, SkillDefinition] = {}
        self._active_skills: Set[str] = set()
        self._compiled_patterns: Dict[str, List[re.Pattern]] = {}

    # ==================== 注册与管理 ====================

    def register(self, skill: SkillDefinition) -> None:
        """注册技能

        Args:
            skill: 技能定义
        """
        self._registry[skill.name] = skill

        # 预编译正则模式
        if skill.patterns:
            self._compiled_patterns[skill.name] = [
                re.compile(p, re.IGNORECASE) for p in skill.patterns
            ]

        # 核心技能自动激活
        if skill.category == SkillCategory.CORE:
            self._active_skills.add(skill.name)

    def unregister(self, skill_name: str) -> bool:
        """注销技能

        Args:
            skill_name: 技能名称

        Returns:
            是否注销成功
        """
        if skill_name not in self._registry:
            return False

        del self._registry[skill_name]
        self._active_skills.discard(skill_name)

        if skill_name in self._compiled_patterns:
            del self._compiled_patterns[skill_name]

        return True

    def get_skill(self, skill_name: str) -> Optional[SkillDefinition]:
        """获取技能定义"""
        return self._registry.get(skill_name)

    def list_skills(self) -> List[SkillDefinition]:
        """列出所有已注册的技能"""
        return list(self._registry.values())

    def list_active_skills(self) -> List[SkillDefinition]:
        """列出当前激活的技能"""
        return [self._registry[name] for name in self._active_skills if name in self._registry]

    # ==================== 意图路由 ====================

    def resolve(
        self,
        user_query: str,
        top_k: int = 3,
        threshold: float = 0.3
    ) -> List[SkillDefinition]:
        """根据用户输入解析需要激活的技能

        路由流程:
        1. 关键词匹配（快速路径）
        2. 正则模式匹配
        3. 语义相似度匹配（可选，需要 embedding）
        4. 依赖解析

        Args:
            user_query: 用户输入
            top_k: 最多返回的技能数量
            threshold: 最低匹配分数阈值

        Returns:
            需要激活的技能列表
        """
        matches: List[IntentMatch] = []

        for skill in self._registry.values():
            # 核心技能始终匹配
            if skill.category == SkillCategory.CORE:
                matches.append(IntentMatch(skill, 1.0, "core"))
                continue

            # 1. 关键词匹配
            keyword_score = self._match_keywords(user_query, skill)
            if keyword_score > 0:
                matches.append(IntentMatch(
                    skill, keyword_score, "keyword",
                    self._find_matched_keyword(user_query, skill.keywords)
                ))
                continue

            # 2. 正则模式匹配
            pattern_match = self._match_patterns(user_query, skill)
            if pattern_match:
                matches.append(IntentMatch(
                    skill, 0.9, "pattern", pattern_match
                ))
                continue

            # 3. 简单语义匹配（基于描述和示例的关键词重叠）
            semantic_score = self._simple_semantic_match(user_query, skill)
            if semantic_score >= threshold:
                matches.append(IntentMatch(skill, semantic_score, "semantic"))

        # 按分数和优先级排序
        matches.sort(key=lambda m: (m.score, m.skill.priority), reverse=True)

        # 取 top_k
        top_matches = matches[:top_k]

        # 收集技能并解析依赖
        result_skills: Set[SkillDefinition] = set()
        for match in top_matches:
            if match.score >= threshold:
                result_skills.add(match.skill)
                # 添加依赖的技能
                for dep_name in match.skill.requires:
                    dep_skill = self._registry.get(dep_name)
                    if dep_skill:
                        result_skills.add(dep_skill)

        # 更新激活状态
        self._update_active_skills(result_skills)

        return list(result_skills)

    def _match_keywords(self, query: str, skill: SkillDefinition) -> float:
        """关键词匹配

        Returns:
            匹配分数 0-1
        """
        query_lower = query.lower()
        matched_count = 0

        for keyword in skill.keywords:
            if keyword.lower() in query_lower:
                matched_count += 1

        if matched_count == 0:
            return 0.0

        # 匹配的关键词越多，分数越高
        return min(1.0, 0.7 + 0.1 * matched_count)

    def _find_matched_keyword(self, query: str, keywords: List[str]) -> str:
        """找到匹配的关键词"""
        query_lower = query.lower()
        for keyword in keywords:
            if keyword.lower() in query_lower:
                return keyword
        return ""

    def _match_patterns(self, query: str, skill: SkillDefinition) -> Optional[str]:
        """正则模式匹配

        Returns:
            匹配的文本，无匹配返回 None
        """
        patterns = self._compiled_patterns.get(skill.name, [])
        for pattern in patterns:
            match = pattern.search(query)
            if match:
                return match.group()
        return None

    def _simple_semantic_match(self, query: str, skill: SkillDefinition) -> float:
        """简单语义匹配（基于词汇重叠）

        这是一个简化版本，不依赖 embedding。
        可以后续升级为真正的向量语义匹配。

        Returns:
            匹配分数 0-1
        """
        # 构建技能的关键词集合（从描述和示例中提取）
        skill_words = set()

        # 从描述中提取（简单分词）
        skill_words.update(self._tokenize(skill.description))

        # 从示例中提取
        for example in skill.examples:
            skill_words.update(self._tokenize(example))

        if not skill_words:
            return 0.0

        # 计算查询词与技能词的重叠度
        query_words = set(self._tokenize(query))
        if not query_words:
            return 0.0

        intersection = query_words & skill_words
        union = query_words | skill_words

        # Jaccard 相似度
        jaccard = len(intersection) / len(union) if union else 0.0

        # 加权：匹配的词在技能描述中更重要
        skill_match_ratio = len(intersection) / len(skill_words) if skill_words else 0.0

        return 0.6 * jaccard + 0.4 * skill_match_ratio

    def _tokenize(self, text: str) -> List[str]:
        """简单分词（支持中英文）"""
        # 移除标点，转小写
        text = re.sub(r'[^\w\s\u4e00-\u9fff]', ' ', text.lower())

        # 英文按空格分词
        words = text.split()

        # 中文按字符分词（简化处理）
        result = []
        for word in words:
            if re.match(r'^[\u4e00-\u9fff]+$', word):
                # 中文：按2-gram分词
                for i in range(len(word)):
                    result.append(word[i])
                    if i < len(word) - 1:
                        result.append(word[i:i+2])
            else:
                result.append(word)

        return result

    def _update_active_skills(self, skills: Set[SkillDefinition]) -> None:
        """更新激活的技能集合"""
        # 保留核心技能
        core_skills = {
            name for name, skill in self._registry.items()
            if skill.category == SkillCategory.CORE
        }

        # 更新激活集合
        self._active_skills = core_skills | {s.name for s in skills}

    # ==================== 工具访问 ====================

    def get_active_tools(self) -> List[Callable]:
        """获取当前激活技能的所有工具

        Returns:
            工具函数列表
        """
        tools = []
        seen = set()  # 避免重复

        for skill_name in self._active_skills:
            skill = self._registry.get(skill_name)
            if skill:
                for tool in skill.tools:
                    tool_name = getattr(tool, 'name', str(tool))
                    if tool_name not in seen:
                        tools.append(tool)
                        seen.add(tool_name)

        return tools

    def get_all_tools(self) -> List[Callable]:
        """获取所有已注册的工具

        Returns:
            所有工具函数列表
        """
        tools = []
        seen = set()

        for skill in self._registry.values():
            for tool in skill.tools:
                tool_name = getattr(tool, 'name', str(tool))
                if tool_name not in seen:
                    tools.append(tool)
                    seen.add(tool_name)

        return tools

    def get_tools_by_skill(self, skill_name: str) -> List[Callable]:
        """获取指定技能的工具

        Args:
            skill_name: 技能名称

        Returns:
            工具函数列表
        """
        skill = self._registry.get(skill_name)
        if skill:
            return skill.tools.copy()
        return []

    # ==================== 手动激活/停用 ====================

    def activate(self, skill_name: str) -> bool:
        """手动激活技能

        Args:
            skill_name: 技能名称

        Returns:
            是否激活成功
        """
        if skill_name not in self._registry:
            return False

        skill = self._registry[skill_name]

        # 检查互斥
        for conflict in skill.conflicts:
            if conflict in self._active_skills:
                return False

        # 激活依赖
        for dep in skill.requires:
            if dep in self._registry:
                self._active_skills.add(dep)

        self._active_skills.add(skill_name)
        return True

    def deactivate(self, skill_name: str) -> bool:
        """手动停用技能

        Args:
            skill_name: 技能名称

        Returns:
            是否停用成功
        """
        if skill_name not in self._registry:
            return False

        skill = self._registry[skill_name]

        # 核心技能不能停用
        if skill.category == SkillCategory.CORE:
            return False

        self._active_skills.discard(skill_name)
        return True

    def reset(self) -> None:
        """重置为初始状态（仅保留核心技能激活）"""
        self._active_skills = {
            name for name, skill in self._registry.items()
            if skill.category == SkillCategory.CORE
        }

    # ==================== 系统提示 ====================

    def get_system_prompt_additions(self) -> str:
        """获取激活技能的额外系统提示

        Returns:
            合并后的系统提示文本
        """
        prompts = []

        for skill_name in self._active_skills:
            skill = self._registry.get(skill_name)
            if skill and skill.system_prompt:
                prompts.append(f"## {skill.display_name}\n{skill.system_prompt}")

        return "\n\n".join(prompts)

    def get_skills_summary(self) -> str:
        """获取技能摘要（用于调试）

        Returns:
            技能摘要文本
        """
        lines = ["=== SkillManager 状态 ===", ""]

        # 按类别分组
        by_category: Dict[SkillCategory, List[SkillDefinition]] = {}
        for skill in self._registry.values():
            if skill.category not in by_category:
                by_category[skill.category] = []
            by_category[skill.category].append(skill)

        for category in [SkillCategory.CORE, SkillCategory.ON_DEMAND, SkillCategory.SYSTEM]:
            skills = by_category.get(category, [])
            if skills:
                lines.append(f"[{category.value.upper()}]")
                for skill in skills:
                    status = "[*]" if skill.name in self._active_skills else "[ ]"
                    tool_count = len(skill.tools)
                    lines.append(f"  {status} {skill.display_name} ({skill.name}) - {tool_count} tools")
                lines.append("")

        return "\n".join(lines)

    def __repr__(self):
        total = len(self._registry)
        active = len(self._active_skills)
        return f"SkillManager(skills={total}, active={active})"
