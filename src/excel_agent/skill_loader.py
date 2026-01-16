"""SkillLoader - 技能加载器

统一的技能加载和管理接口，整合 SkillScanner 和 SkillManager。
实现 Claude Code 风格的技能系统：
1. 启动时扫描技能文件，只加载元数据
2. 生成紧凑的技能列表供系统提示使用
3. 需要时才加载完整技能内容
"""

import logging
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Set

from .skill_manager import SkillCategory, SkillDefinition, SkillManager
from .skill_scanner import SkillMetadata, SkillScanner, get_skill_scanner


# 配置日志
logger = logging.getLogger("excel_agent.skills")


@dataclass
class SkillMatchResult:
    """技能匹配结果（用于日志）"""
    skill_name: str
    display_name: str
    score: float
    match_type: str  # "core", "keyword", "pattern", "semantic"
    matched_text: str = ""

    def __str__(self):
        if self.match_type == "core":
            return f"  [CORE] {self.display_name} (始终激活)"
        elif self.match_type == "keyword":
            return f"  [关键词] {self.display_name} (score={self.score:.2f}, 匹配: '{self.matched_text}')"
        elif self.match_type == "pattern":
            return f"  [正则] {self.display_name} (score={self.score:.2f}, 匹配: '{self.matched_text}')"
        else:
            return f"  [语义] {self.display_name} (score={self.score:.2f})"


class SkillLoader:
    """技能加载器

    提供统一的接口来：
    1. 获取技能列表（用于系统提示，节省 token）
    2. 按需加载技能工具
    3. 管理技能激活状态

    使用示例:
    ```python
    loader = SkillLoader(tools_registry)
    loader.initialize()

    # 获取技能列表（节省 token）
    skill_list = loader.get_skill_list_for_prompt()

    # 根据用户输入激活技能
    activated = loader.activate_skills_for_query("帮我筛选数据")

    # 获取激活的工具
    tools = loader.get_active_tools()
    ```
    """

    def __init__(self, tools_registry: Dict[str, Callable] = None):
        """初始化技能加载器

        Args:
            tools_registry: 工具注册表 {tool_name: tool_function}
        """
        self._tools_registry = tools_registry or {}
        self._scanner: Optional[SkillScanner] = None
        self._manager: Optional[SkillManager] = None
        self._initialized = False

    def initialize(self, skills_dir: str = None) -> int:
        """初始化：扫描技能文件

        Args:
            skills_dir: 技能文件目录，默认使用模块同级 skills/ 目录

        Returns:
            发现的技能数量
        """
        if skills_dir:
            self._scanner = SkillScanner(skills_dir)
        else:
            self._scanner = get_skill_scanner()

        count = self._scanner.scan()
        self._manager = SkillManager()
        self._initialized = True

        # 日志：输出扫描到的技能列表
        self._log_scanned_skills()

        # 预加载核心技能到 Manager
        core_count = 0
        for metadata in self._scanner.get_all_metadata():
            if metadata.category == SkillCategory.CORE:
                self._load_skill_to_manager(metadata.name)
                core_count += 1

        logger.info(f"[Skills] 预加载 {core_count} 个核心技能")

        return count

    def _log_scanned_skills(self):
        """输出扫描到的技能列表日志"""
        all_metadata = self._scanner.get_all_metadata()

        # 按类别分组
        by_category: Dict[SkillCategory, List[SkillMetadata]] = {}
        for m in all_metadata:
            if m.category not in by_category:
                by_category[m.category] = []
            by_category[m.category].append(m)

        print("\n" + "=" * 60)
        print("📦 Skills 扫描完成")
        print("=" * 60)

        for category in [SkillCategory.CORE, SkillCategory.ON_DEMAND, SkillCategory.SYSTEM]:
            skills = by_category.get(category, [])
            if skills:
                category_names = {
                    SkillCategory.CORE: "🔵 核心技能 (始终激活)",
                    SkillCategory.ON_DEMAND: "🟢 按需技能",
                    SkillCategory.SYSTEM: "⚙️  系统技能"
                }
                print(f"\n{category_names[category]}:")
                for skill in sorted(skills, key=lambda x: -x.priority):
                    tool_count = len(skill.tool_names)
                    keyword_count = len(skill.keywords)
                    print(f"  - {skill.display_name} ({skill.name})")
                    print(f"    📝 {skill.description[:40]}...")
                    print(f"    🔧 {tool_count} 工具, 🏷️  {keyword_count} 关键词, ⭐ 优先级 {skill.priority}")

        print("\n" + "=" * 60)
        print(f"📊 总计: {len(all_metadata)} 个技能")
        print("=" * 60 + "\n")

    def _load_skill_to_manager(self, skill_name: str) -> bool:
        """将技能加载到 Manager"""
        if not self._scanner:
            return False

        skill_def = self._scanner.to_skill_definition(skill_name, self._tools_registry)
        if skill_def:
            self._manager.register(skill_def)
            return True
        return False

    def get_skill_list_for_prompt(self) -> str:
        """获取技能列表字符串（用于系统提示，节省 token）

        Returns:
            格式化的技能列表
        """
        if not self._scanner:
            return "暂无可用技能。"

        return self._scanner.get_skill_list_prompt()

    def activate_skills_for_query(
        self,
        user_query: str,
        top_k: int = 3,
        threshold: float = 0.3
    ) -> List[str]:
        """根据用户查询激活相关技能

        Args:
            user_query: 用户输入
            top_k: 最多激活的技能数量
            threshold: 最低匹配阈值

        Returns:
            激活的技能名称列表
        """
        if not self._scanner or not self._manager:
            return []

        # 使用 Scanner 的元数据进行匹配（带日志）
        matched_skills, match_results = self._match_skills_with_log(user_query, top_k, threshold)

        # 输出匹配日志
        self._log_skill_matching(user_query, match_results, threshold)

        # 加载匹配的技能到 Manager（懒加载）
        activated = []
        for skill_name in matched_skills:
            if skill_name not in [s.name for s in self._manager.list_skills()]:
                self._load_skill_to_manager(skill_name)
                logger.debug(f"[Skills] 懒加载技能: {skill_name}")

            if self._manager.activate(skill_name):
                activated.append(skill_name)

        # 加载依赖的技能
        deps_loaded = []
        for skill_name in list(activated):
            metadata = self._scanner.get_metadata(skill_name)
            if metadata:
                content = self._scanner.load_full_skill(skill_name)
                if content:
                    for dep_name in content.requires:
                        if dep_name not in activated:
                            self._load_skill_to_manager(dep_name)
                            self._manager.activate(dep_name)
                            activated.append(dep_name)
                            deps_loaded.append(dep_name)

        if deps_loaded:
            print(f"  📎 加载依赖技能: {', '.join(deps_loaded)}")

        return activated

    def _match_skills_with_log(
        self,
        user_query: str,
        top_k: int,
        threshold: float
    ) -> tuple[List[str], List[SkillMatchResult]]:
        """基于元数据匹配技能（带详细日志）

        Returns:
            (匹配的技能名称列表, 匹配结果详情列表)
        """
        import re

        match_results: List[SkillMatchResult] = []
        query_lower = user_query.lower()

        for metadata in self._scanner.get_all_metadata():
            # 核心技能始终包含
            if metadata.category == SkillCategory.CORE:
                match_results.append(SkillMatchResult(
                    skill_name=metadata.name,
                    display_name=metadata.display_name,
                    score=1.0,
                    match_type="core"
                ))
                continue

            score = 0.0
            match_type = ""
            matched_text = ""

            # 1. 关键词匹配
            matched_keywords = []
            for keyword in metadata.keywords:
                if keyword.lower() in query_lower:
                    matched_keywords.append(keyword)

            if matched_keywords:
                score = max(score, 0.7 + 0.1 * min(len(matched_keywords), 3))
                match_type = "keyword"
                matched_text = ", ".join(matched_keywords[:3])

            # 2. 正则模式匹配
            if score < 0.9:  # 如果关键词没有达到高分，尝试正则
                for pattern in metadata.patterns:
                    try:
                        match = re.search(pattern, user_query, re.IGNORECASE)
                        if match:
                            score = max(score, 0.9)
                            match_type = "pattern"
                            matched_text = match.group()
                            break
                    except re.error:
                        pass

            # 3. 描述词匹配（简单语义）
            if score == 0:
                desc_words = set(metadata.description.lower().split())
                query_words = set(query_lower.split())
                overlap = desc_words & query_words
                if overlap:
                    score = 0.3 + 0.1 * len(overlap)
                    match_type = "semantic"
                    matched_text = ", ".join(list(overlap)[:3])

            if score >= threshold:
                match_results.append(SkillMatchResult(
                    skill_name=metadata.name,
                    display_name=metadata.display_name,
                    score=score,
                    match_type=match_type,
                    matched_text=matched_text
                ))

        # 按分数排序
        match_results.sort(key=lambda x: (-x.score, -self._get_priority(x.skill_name)))

        # 返回 top_k
        top_results = match_results[:top_k]
        return [r.skill_name for r in top_results], match_results

    def _log_skill_matching(
        self,
        user_query: str,
        match_results: List[SkillMatchResult],
        threshold: float
    ):
        """输出技能匹配日志"""
        print("\n" + "-" * 50)
        print(f"🔍 技能匹配 | 查询: \"{user_query[:50]}{'...' if len(user_query) > 50 else ''}\"")
        print("-" * 50)

        if not match_results:
            print("  ⚠️  未匹配到任何技能")
        else:
            # 分组显示
            above_threshold = [r for r in match_results if r.score >= threshold]
            below_threshold = [r for r in match_results if r.score < threshold and r.match_type != "core"]

            print(f"\n✅ 激活的技能 (score >= {threshold}):")
            for result in above_threshold:
                print(str(result))

            if below_threshold:
                print(f"\n⏸️  未激活 (score < {threshold}):")
                for result in below_threshold[:3]:  # 只显示前3个
                    print(f"  [-] {result.display_name} (score={result.score:.2f})")

        print("-" * 50 + "\n")

    def _get_priority(self, skill_name: str) -> int:
        """获取技能优先级"""
        metadata = self._scanner.get_metadata(skill_name)
        return metadata.priority if metadata else 0

    def get_active_tools(self) -> List[Callable]:
        """获取当前激活的工具列表

        Returns:
            工具函数列表
        """
        if not self._manager:
            return []
        return self._manager.get_active_tools()

    def get_all_tools(self) -> List[Callable]:
        """获取所有已注册的工具

        Returns:
            所有工具函数列表
        """
        if not self._manager:
            return list(self._tools_registry.values())
        return self._manager.get_all_tools()

    def get_system_prompt_additions(self) -> str:
        """获取激活技能的系统提示补充

        Returns:
            系统提示补充文本
        """
        if not self._manager:
            return ""
        return self._manager.get_system_prompt_additions()

    def get_active_skill_names(self) -> List[str]:
        """获取当前激活的技能名称列表"""
        if not self._manager:
            return []
        return [s.display_name for s in self._manager.list_active_skills()]

    def get_skills_summary(self) -> Dict[str, any]:
        """获取技能系统摘要信息

        Returns:
            包含技能统计信息的字典
        """
        if not self._scanner:
            return {"total": 0, "core": 0, "on_demand": 0, "system": 0}

        all_metadata = self._scanner.get_all_metadata()
        active_skills = self._manager.list_active_skills() if self._manager else []

        by_category = {}
        for m in all_metadata:
            cat = m.category.value
            if cat not in by_category:
                by_category[cat] = []
            by_category[cat].append({
                "name": m.name,
                "display_name": m.display_name,
                "tools": len(m.tool_names),
                "keywords": len(m.keywords)
            })

        return {
            "total": len(all_metadata),
            "core": len(by_category.get("core", [])),
            "on_demand": len(by_category.get("on_demand", [])),
            "system": len(by_category.get("system", [])),
            "active_count": len(active_skills),
            "active_names": [s.display_name for s in active_skills],
            "by_category": by_category
        }

    def log_prompt_context(self) -> str:
        """获取并打印用于系统提示的技能列表

        Returns:
            技能列表字符串
        """
        skill_list = self.get_skill_list_for_prompt()

        print("\n" + "=" * 50)
        print("📋 系统提示中的技能列表 (节省 Token)")
        print("=" * 50)
        print(skill_list)
        print("=" * 50 + "\n")

        return skill_list

    def reset(self) -> None:
        """重置激活状态"""
        if self._manager:
            self._manager.reset()


# ==================== 工具注册表构建 ====================

def build_tools_registry() -> Dict[str, Callable]:
    """构建工具注册表

    Returns:
        {tool_name: tool_function} 映射
    """
    from .tools import ALL_TOOLS

    registry = {}

    # 直接使用 ALL_TOOLS 列表
    for tool in ALL_TOOLS:
        # LangChain StructuredTool 的 name 属性是工具名称
        if hasattr(tool, 'name'):
            registry[tool.name] = tool

    # 输出调试日志
    print(f"\n[Skills] 工具注册表构建完成: {len(registry)} 个工具")
    if registry:
        tool_names = sorted(registry.keys())
        print(f"[Skills] 已注册工具: {', '.join(tool_names[:10])}{'...' if len(tool_names) > 10 else ''}")

    return registry


# ==================== 全局实例 ====================

_skill_loader: Optional[SkillLoader] = None


def get_skill_loader() -> SkillLoader:
    """获取全局 SkillLoader 实例"""
    global _skill_loader
    if _skill_loader is None:
        registry = build_tools_registry()
        _skill_loader = SkillLoader(registry)
        _skill_loader.initialize()
    return _skill_loader


def reset_skill_loader() -> None:
    """重置全局 SkillLoader 实例"""
    global _skill_loader
    _skill_loader = None
