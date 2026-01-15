# ExcelMind v2.0 技术蓝图

> 本文档记录 ExcelMind 下一代架构（v2.0）的技术设计，聚焦于 **写入能力**、**Skills 动态路由**、**代码解释器** 和 **性能优化**。

---

## 一、v1.x 回顾与现状

### 1.1 已实现能力（v1.0 - v1.3）

| 版本 | 核心能力 |
|------|---------|
| v1.0 | 基础分析、自然语言查询、ECharts 可视化、本地知识库 |
| v1.1 | LangChain ReAct Agent、真正流式输出、可折叠思考过程 |
| v1.2 | Anthropic 原生支持、模型容错降级、多表管理、智能联表 |
| v1.3 | 工作表切换、递归限制修复、演示视频集成 |

### 1.2 当前架构

```
┌─────────────────────────────────────────────────────────────────┐
│                    ExcelMind v1.3 架构                          │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌──────────────┐     ┌──────────────┐     ┌──────────────┐     │
│  │   FastAPI    │ ──> │  ReAct Agent │ ──> │    Tools     │     │
│  │   (api.py)   │     │  (stream.py) │     │  (tools.py)  │     │
│  └──────────────┘     └──────────────┘     └──────────────┘     │
│         │                    │                    │             │
│         v                    v                    v             │
│  ┌──────────────┐     ┌──────────────┐     ┌──────────────┐     │
│  │ MultiExcel   │     │  KnowledgeDB │     │   pandas     │     │
│  │   Loader     │     │   (Chroma)   │     │  DataFrame   │     │
│  └──────────────┘     └──────────────┘     └──────────────┘     │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 1.3 当前局限

| 局限 | 影响 |
|------|------|
| **只读不写** | 无法修改 Excel 内容、添加公式、设置格式 |
| **静态工具列表** | 所有工具始终加载，占用上下文 Token |
| **无缓存机制** | 重复查询消耗计算资源 |
| **公式计算依赖外部** | pandas 只读值，无法执行公式 |
| **复杂逻辑受限** | 工具固定，无法处理任意数据操作 |

---

## 二、v2.0 目标架构

### 2.1 架构全景

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        ExcelMind v2.0 架构                               │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  ┌───────────────────────────────────────────────────────────────────┐  │
│  │                      Intent Router (意图路由)                       │  │
│  │   用户消息 → 关键词匹配 / 语义检索 → 动态激活 Skills               │  │
│  └───────────────────────────────────────────────────────────────────┘  │
│                                    │                                    │
│           ┌────────────────────────┼────────────────────────┐           │
│           v                        v                        v           │
│  ┌─────────────────┐   ┌─────────────────┐   ┌─────────────────┐        │
│  │   Core Skills   │   │ On-Demand Skills │   │  System Skills  │        │
│  │  (始终激活)      │   │   (按需加载)      │   │   (流程控制)    │        │
│  │ • query_data    │   │ • Modification   │   │ • file_io       │        │
│  │ • project_info  │   │ • Formatting     │   │ • code_exec     │        │
│  │                 │   │ • Analytics      │   │ • batch_process │        │
│  │                 │   │ • Visualization  │   │                 │        │
│  └─────────────────┘   └─────────────────┘   └─────────────────┘        │
│                                    │                                    │
│  ┌───────────────────────────────────────────────────────────────────┐  │
│  │                       Dual Engine (双引擎)                         │  │
│  │   ┌──────────────────┐           ┌──────────────────┐             │  │
│  │   │     pandas       │   <───>   │    openpyxl      │             │  │
│  │   │   (分析引擎)      │    同步    │   (操作引擎)      │             │  │
│  │   │  • 快速查询       │           │  • 读写公式       │             │  │
│  │   │  • 聚合统计       │           │  • 修改单元格     │             │  │
│  │   │  • 分组计算       │           │  • 格式样式       │             │  │
│  │   └──────────────────┘           └──────────────────┘             │  │
│  └───────────────────────────────────────────────────────────────────┘  │
│                                    │                                    │
│  ┌───────────────────────────────────────────────────────────────────┐  │
│  │                     Cache Layer (缓存层)                           │  │
│  │   • 查询结果缓存 (LRU)  • 热数据预加载  • 增量同步                  │  │
│  └───────────────────────────────────────────────────────────────────┘  │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

### 2.2 核心升级点

| 升级点 | 说明 |
|--------|------|
| **双引擎架构** | pandas + openpyxl 各司其职，实现读写分离 |
| **Skills 动态路由** | 根据意图按需加载工具，减少 Token 消耗 |
| **写入能力** | 支持修改单元格、添加公式、设置格式 |
| **代码解释器** | 沙箱执行任意 Python 代码，解决复杂逻辑 |
| **缓存层** | 查询结果缓存，提升响应速度 |

---

## 三、双引擎架构设计

### 3.1 引擎职责划分

| 引擎 | 职责 | 优势 | 劣势 |
|------|------|------|------|
| **pandas** | 查询、筛选、聚合、统计 | 快速（10-100x）、语法简洁 | 只读、丢失公式/格式 |
| **openpyxl** | 写入、公式、格式、保存 | 完整 Excel 支持 | 查询慢、需手写循环 |

### 3.2 ExcelDocument 类设计

```python
class ExcelDocument:
    """双引擎 Excel 文档管理器"""

    # 核心属性
    file_path: str
    workbook: openpyxl.Workbook      # 操作引擎
    dataframe: pd.DataFrame           # 分析引擎
    active_sheet: str
    is_dirty: bool                    # 是否有未保存修改
    change_log: List[Change]          # 变更记录

    # 加载与保存
    def load(path: str, sheet_name: str = None) -> None
    def save(path: str = None) -> None
    def save_as(path: str) -> None

    # 引擎访问
    def get_read_engine(self) -> pd.DataFrame
    def get_write_engine(self) -> openpyxl.Worksheet

    # 数据同步
    def sync_workbook_to_df(self) -> None      # openpyxl -> pandas
    def sync_df_to_workbook(self) -> None      # pandas -> openpyxl (谨慎使用)

    # 单元格操作
    def read_cell(row: int, col: int) -> CellValue
    def write_cell(row: int, col: int, value: Any) -> None
    def read_formula(row: int, col: int) -> str
    def write_formula(row: int, col: int, formula: str) -> None

    # 范围操作
    def read_range(start: str, end: str) -> List[List[Any]]
    def write_range(start: str, data: List[List[Any]]) -> None

    # 行列操作
    def insert_rows(row: int, count: int = 1) -> None
    def delete_rows(row: int, count: int = 1) -> None
    def insert_cols(col: int, count: int = 1) -> None
    def delete_cols(col: int, count: int = 1) -> None

    # 事务支持
    def begin_transaction(self) -> None
    def commit(self) -> None
    def rollback(self) -> None
```

### 3.3 数据同步策略

```
┌─────────────────────────────────────────────────────────────┐
│                    Lazy Sync (惰性同步)                      │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  读取操作 (优先使用 pandas):                                │
│  ┌─────────────┐     ┌─────────────┐                        │
│  │ 查询请求    │ ──> │  DataFrame  │ ──> 返回结果           │
│  └─────────────┘     └─────────────┘                        │
│                                                             │
│  写入操作 (直接操作 openpyxl):                              │
│  ┌─────────────┐     ┌─────────────┐     ┌─────────────┐    │
│  │ 写入请求    │ ──> │  Workbook   │ ──> │ 标记 dirty  │    │
│  └─────────────┘     └─────────────┘     └─────────────┘    │
│                                                             │
│  同步触发条件:                                              │
│  • 写入后的下一次读取 (按需同步)                            │
│  • 显式调用 sync_workbook_to_df()                          │
│  • 保存文件前                                               │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 四、Skills 动态路由架构

### 4.1 Skill 定义规范

```python
@dataclass
class SkillDefinition:
    """技能定义"""
    name: str                    # 唯一标识: "data_modification"
    display_name: str            # 显示名称: "数据修改"
    description: str             # 语义描述 (用于检索)
    tools: List[Callable]        # 包含的工具函数
    examples: List[str]          # 触发示例

    # 激活配置
    keywords: List[str]          # 强触发关键词
    threshold: float = 0.7       # 语义相似度阈值
    priority: int = 0            # 优先级 (冲突时使用)

    # 依赖关系
    requires: List[str] = []     # 前置 Skills
    conflicts: List[str] = []    # 互斥 Skills
```

### 4.2 技能分类

```
┌─────────────────────────────────────────────────────────────────┐
│                        Skills Registry                          │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  1. Core Skills (核心技能 - 始终激活)                           │
│     ├── query_data        多态查询 (filter/agg/group/search)    │
│     └── project_info      元数据获取、结构信息                  │
│                                                                 │
│  2. On-Demand Skills (按需加载 - 意图触发)                      │
│     ├── Modification      写入、更新、删除                      │
│     │   └── Tools: write_cell, write_range, delete_rows...      │
│     ├── Formula           公式读写与管理                        │
│     │   └── Tools: read_formula, write_formula, list_formulas   │
│     ├── Formatting        格式、样式、条件格式                  │
│     │   └── Tools: set_style, set_border, conditional_format    │
│     ├── Analytics         高级分析、透视表                      │
│     │   └── Tools: pivot_table, correlation, regression         │
│     └── Visualization     图表生成与配置                        │
│         └── Tools: create_chart, update_chart, export_chart     │
│                                                                 │
│  3. System Skills (系统技能 - 流程控制)                         │
│     ├── file_io           保存、另存、导出                      │
│     ├── code_exec         Python 代码执行 (沙箱)                │
│     └── batch_process     批量处理任务                          │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 4.3 SkillManager 设计

```python
class SkillManager:
    """技能管理器"""

    registry: Dict[str, SkillDefinition]   # 技能注册表
    active_skills: Set[str]                 # 当前激活的技能
    embedding_model: EmbeddingModel         # 语义检索模型

    # 注册与管理
    def register(skill: SkillDefinition) -> None
    def unregister(skill_name: str) -> None
    def list_skills() -> List[SkillDefinition]

    # 意图路由
    def resolve(user_query: str) -> List[SkillDefinition]:
        """
        根据用户输入解析需要激活的技能
        1. 关键词匹配 (快速路径)
        2. 语义相似度 (泛化路径)
        3. 依赖解析 (自动加载前置技能)
        """

    # 工具注入
    def get_active_tools() -> List[Callable]:
        """返回当前激活技能的所有工具"""

    def get_system_prompt_additions() -> str:
        """返回激活技能的额外系统提示"""
```

### 4.4 意图路由流程

```
用户输入: "帮我把A列的数据求和，然后写入到B1单元格"
                    │
                    v
┌─────────────────────────────────────────────────────────────┐
│                   Intent Router                              │
├─────────────────────────────────────────────────────────────┤
│  Step 1: 关键词匹配                                          │
│     "写入" → matches Modification Skill                      │
│     "求和" → matches Core Query Skill                        │
│                                                             │
│  Step 2: 语义增强 (可选)                                     │
│     embedding("帮我把A列的数据求和，然后写入到B1单元格")     │
│     → similarity_search(skill_descriptions)                  │
│     → [Modification: 0.85, Formula: 0.72, ...]              │
│                                                             │
│  Step 3: 依赖解析                                            │
│     Modification requires Core (✓ 已激活)                    │
│                                                             │
│  Step 4: 激活技能                                            │
│     active_skills = {Core, Modification}                     │
│     active_tools = [query_data, write_cell, write_range...]  │
└─────────────────────────────────────────────────────────────┘
```

---

## 五、写入能力设计

### 5.1 能力边界

| 能力 | 支持级别 | 说明 |
|------|---------|------|
| **单元格写入** | 完全支持 | 写入值、数字、日期等 |
| **范围写入** | 完全支持 | 批量写入数据区域 |
| **公式写入** | 完全支持 | 写入 Excel 公式文本 |
| **公式读取** | 完全支持 | 获取单元格公式 |
| **行列操作** | 完全支持 | 插入、删除行/列 |
| **基础格式** | 完全支持 | 字体、颜色、对齐、边框 |
| **条件格式** | 有限支持 | 基础条件格式规则 |
| **图表写入** | 有限支持 | 基础图表类型 |
| **公式计算** | 不支持 | 需保存后用 Excel 打开 |
| **VBA 宏** | 不支持 | openpyxl 不执行代码 |

### 5.2 写入工具定义

```python
# Modification Skill 工具集

@tool
def write_cell(
    cell: str,           # "A1", "B2" 等
    value: Any,          # 写入的值
    sheet: str = None    # 目标工作表
) -> Dict:
    """写入单个单元格"""

@tool
def write_range(
    start_cell: str,     # 起始单元格 "A1"
    data: List[List],    # 二维数据数组
    sheet: str = None
) -> Dict:
    """批量写入数据区域"""

@tool
def write_formula(
    cell: str,           # 目标单元格
    formula: str,        # Excel 公式 "=SUM(A1:A10)"
    sheet: str = None
) -> Dict:
    """写入公式"""

@tool
def insert_rows(
    row: int,            # 在此行之前插入
    count: int = 1,      # 插入行数
    sheet: str = None
) -> Dict:
    """插入行"""

@tool
def delete_rows(
    start_row: int,      # 起始行
    end_row: int = None, # 结束行 (含)
    sheet: str = None
) -> Dict:
    """删除行"""

@tool
def save_file(
    path: str = None     # 另存为路径 (空则覆盖原文件)
) -> Dict:
    """保存文件"""
```

### 5.3 格式化工具定义

```python
# Formatting Skill 工具集

@tool
def set_cell_style(
    cell: str,                    # 单元格或范围 "A1" 或 "A1:B10"
    font: Dict = None,            # {"bold": True, "size": 12, "color": "FF0000"}
    fill: Dict = None,            # {"color": "FFFF00"}
    alignment: Dict = None,       # {"horizontal": "center", "vertical": "center"}
    border: Dict = None,          # {"style": "thin", "color": "000000"}
    number_format: str = None     # "#,##0.00" 或 "0%"
) -> Dict:
    """设置单元格样式"""

@tool
def auto_fit_columns(
    columns: List[str] = None,    # 列列表 ["A", "B"] 或 None (全部)
    sheet: str = None
) -> Dict:
    """自动调整列宽"""

@tool
def merge_cells(
    range: str,                   # "A1:C1"
    sheet: str = None
) -> Dict:
    """合并单元格"""

@tool
def add_conditional_format(
    range: str,                   # 应用范围 "A1:A100"
    rule_type: str,               # "greater_than", "less_than", "between", "duplicate"
    conditions: Dict,             # {"value": 100} 或 {"min": 0, "max": 100}
    style: Dict                   # 匹配时的样式
) -> Dict:
    """添加条件格式"""
```

---

## 六、代码解释器设计

### 6.1 设计目标

- **灵活性**: 执行任意 Python 代码，解决工具无法覆盖的复杂逻辑
- **安全性**: 沙箱隔离，限制危险操作
- **交互性**: 支持多轮代码执行，保持上下文

### 6.2 沙箱架构

```
┌─────────────────────────────────────────────────────────────┐
│                    Code Interpreter Sandbox                  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  允许的模块:                                                │
│  ├── pandas, numpy, scipy                                   │
│  ├── openpyxl (受限访问)                                    │
│  ├── datetime, math, statistics                             │
│  └── json, re, collections                                  │
│                                                             │
│  禁止的操作:                                                │
│  ├── 文件系统访问 (除了工作目录)                            │
│  ├── 网络请求                                               │
│  ├── 系统命令执行                                           │
│  ├── 进程操作                                               │
│  └── 动态代码执行 (eval/exec 嵌套)                          │
│                                                             │
│  资源限制:                                                  │
│  ├── 执行时间: 30秒                                         │
│  ├── 内存限制: 512MB                                        │
│  └── 输出限制: 100KB                                        │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### 6.3 代码执行工具

```python
@tool
def execute_code(
    code: str,                    # Python 代码
    description: str = ""         # 代码目的描述 (用于审计)
) -> Dict:
    """
    在沙箱中执行 Python 代码

    预置变量:
    - df: 当前活跃表的 DataFrame
    - wb: 当前工作簿的 openpyxl Workbook
    - ws: 当前活跃工作表

    返回:
    - result: 代码执行结果
    - output: print 输出
    - variables: 新创建的变量
    """
```

### 6.4 使用示例

```python
# 用户: "帮我找出所有金额异常的记录，定义为超过均值3个标准差的数据"

# LLM 生成的代码:
execute_code("""
import numpy as np

# 计算统计值
mean = df['金额'].mean()
std = df['金额'].std()
threshold = 3 * std

# 找出异常记录
anomalies = df[abs(df['金额'] - mean) > threshold]

# 返回结果
result = {
    'count': len(anomalies),
    'mean': mean,
    'std': std,
    'threshold': threshold,
    'records': anomalies.to_dict('records')[:10]  # 最多返回10条
}
""", description="查找金额异常记录")
```

---

## 七、缓存层设计

### 7.1 缓存策略

```
┌─────────────────────────────────────────────────────────────┐
│                      Cache Architecture                      │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  L1 Cache (内存 - LRU)                                      │
│  ├── 查询结果缓存                                           │
│  ├── 聚合计算结果                                           │
│  └── 有效期: 会话内 / 直到数据修改                          │
│                                                             │
│  L2 Cache (可选 - Redis)                                    │
│  ├── 跨会话缓存                                             │
│  ├── 大文件元数据                                           │
│  └── 有效期: 配置 TTL                                       │
│                                                             │
│  缓存失效触发:                                              │
│  ├── 写入操作后                                             │
│  ├── 文件重新加载                                           │
│  └── 手动清除                                               │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### 7.2 缓存键设计

```python
# 查询缓存键
cache_key = hash(
    table_id,           # 表标识
    sheet_name,         # 工作表
    operation,          # 操作类型
    params_hash,        # 参数哈希
    data_version        # 数据版本号 (写入后递增)
)
```

---

## 八、API 扩展设计

### 8.1 新增接口

```yaml
# 写入操作
POST /write/cell           # 写入单元格
POST /write/range          # 批量写入
POST /write/formula        # 写入公式

# 格式操作
POST /format/style         # 设置样式
POST /format/merge         # 合并单元格
POST /format/conditional   # 条件格式

# 文件操作
POST /file/save            # 保存
POST /file/save-as         # 另存为
POST /file/export          # 导出 (CSV, JSON, PDF)

# 代码执行
POST /code/execute         # 执行代码
GET  /code/history         # 执行历史

# Skills 管理
GET  /skills               # 获取所有技能
GET  /skills/active        # 获取当前激活技能
POST /skills/activate      # 手动激活技能
POST /skills/deactivate    # 手动停用技能
```

### 8.2 WebSocket 实时同步

```yaml
# 支持多客户端实时协作
WS /ws/sync
  Events:
    - cell_changed       # 单元格变更
    - range_changed      # 范围变更
    - sheet_changed      # 工作表变更
    - file_saved         # 文件保存
```

---

## 九、版本规划

### v2.0-alpha (Foundation)

- [ ] **双引擎架构**: ExcelDocument 类实现
- [ ] **基础写入**: write_cell, write_range, save
- [ ] **公式支持**: read_formula, write_formula
- [ ] **行列操作**: insert/delete rows/cols

### v2.0-beta (Skills)

- [ ] **SkillManager**: 技能注册与管理
- [ ] **意图路由**: 关键词 + 语义检索
- [ ] **动态加载**: 按需激活/停用技能
- [ ] **工具合并**: 将现有工具按技能分组

### v2.0-rc (Enhancement)

- [ ] **格式化 Skill**: 完整样式支持
- [ ] **代码解释器**: 沙箱环境实现
- [ ] **缓存层**: LRU 缓存实现
- [ ] **API 扩展**: 新增接口

### v2.0 GA (Production)

- [ ] **性能优化**: 大文件支持、增量加载
- [ ] **文档完善**: API 文档、使用指南
- [ ] **测试覆盖**: 单元测试、集成测试
- [ ] **错误处理**: 完善的错误提示与恢复

### v2.1+ (Future)

- [ ] **原生图表写入**: Excel 内嵌图表
- [ ] **透视表支持**: 创建和修改透视表
- [ ] **混合运行时**: xlwings/win32com 可选后端
- [ ] **多人协作**: WebSocket 实时同步
- [ ] **插件系统**: 第三方技能扩展

---

## 十、迁移指南

### 从 v1.x 迁移

1. **配置文件**: 新增 `skills` 和 `cache` 配置节
2. **API 变更**:
   - `/chat/stream` 保持兼容
   - 新增写入相关接口
3. **工具变更**:
   - 现有工具保持兼容
   - 按技能分组重新组织
4. **数据格式**:
   - ExcelLoader → ExcelDocument
   - 新增 Workbook 属性

---

*文档创建时间: 2025-01-15*
*版本: v2.0 Blueprint*
