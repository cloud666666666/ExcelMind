# 📊 ExcelMind - Excel 数据智能分析助手

> 基于 LangGraph 的智能化 Excel 数据分析工具，让自然语言成为你的数据分析利器！

## 📋 项目简介

ExcelMind 是一个专为 Excel 数据分析设计的 AI 助手，能够理解自然语言并智能分析数据。支持多轮对话、流式输出、ECharts 图表可视化和完整的思考过程展示，让数据分析变得简单直观。

> 🔱 **Fork 声明**：本项目基于 [Gen-Future/ExcelMind](https://github.com/Gen-Future/ExcelMind) 进行二次开发，新增了 **Claude 系列模型原生适配** 以及 **模型调用失败的容错降级机制**。

![Python](https://img.shields.io/badge/Python-3.11+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![FastAPI](https://img.shields.io/badge/FastAPI-Latest-009688)
![LangChain](https://img.shields.io/badge/LangChain-Latest-orange)

## 🎬 演示视频

[![ExcelMind 演示视频](https://img.shields.io/badge/📺_点击观看-Bilibili视频-00A1D6?style=for-the-badge&logo=bilibili)](https://www.bilibili.com/video/BV1VKvWBgEF1/)

> 视频展示了 ExcelMind 的核心功能：自然语言查询、智能联表、知识库检索等

## ✨ 主要功能

### 🎯 核心功能

- **自然语言查询** - 用中文直接提问，无需编写代码或公式
- **多轮对话** - 支持上下文关联的连续追问（如"和上个月相比呢？"）
- **流式输出** - 实时显示 AI 思考过程和回答，响应更流畅
- **智能工具调用** - 自动选择合适的数据分析工具，展示完整推理链路

### 🛠️ 数据分析工具

- **数据筛选** - 多条件筛选、排序、指定返回列
- **聚合统计** - 支持先筛选再聚合
- **分组聚合** - 灵活的分组统计功能
- **关键词搜索** - 可限制搜索范围
- **列统计** - 支持筛选后统计
- **唯一值获取** - 支持筛选后获取
- **数据预览** - 快速查看数据结构
- **时间获取** - 处理相对时间查询
- **数学计算** - 批量精确计算
- **图表生成** - ECharts 可视化，AI 自动推荐图表类型

### 🔧 高级特性

- **多表协同** - 同时管理多个 Excel 表格，支持智能联表
- **本地知识库** - 基于 Chroma 向量数据库的私有知识存储
- **双主题模式** - 亮色/暗色主题一键切换
- **意图过滤** - 自动拒绝与数据无关的闲聊

## 📁 文件结构

```
ExcelMind/
├── config.yaml              # 配置文件（需从 config.example.yaml 复制）
├── config.example.yaml      # 配置文件模板
├── pyproject.toml           # 项目依赖
├── README.md                # 使用说明
├── LICENSE                  # MIT许可证
├── knowledge/               # 知识库文件目录
│   └── *.md                 # Markdown 格式知识文件
├── .vector_db/              # Chroma 向量数据库（自动生成）
├── docs/                    # 文档目录
│   └── card.png             # 社区交流名片
└── src/
    └── excel_agent/
        ├── __init__.py
        ├── main.py          # 程序入口
        ├── api.py           # FastAPI 接口
        ├── config.py        # 配置管理
        ├── excel_loader.py  # Excel 加载器
        ├── graph.py         # LangGraph 工作流
        ├── knowledge_base.py # 知识库管理
        ├── prompts.py       # 提示词模板
        ├── stream.py        # 流式对话核心
        ├── tools.py         # 数据分析工具
        └── frontend/
            └── index.html   # Web 界面
```

## 🚀 快速开始

### 系统要求

- **Python版本**：3.11+
- **包管理器**：[uv](https://github.com/astral-sh/uv)（推荐）或 pip
- **API Key**：OpenAI 兼容的 API Key

### 安装步骤

1. **克隆项目**：

   ```bash
   git clone https://github.com/cloud666666666/ExcelMind.git
   cd ExcelMind
   ```
2. **安装依赖**：

   ```bash
   # 使用 uv（推荐）
   uv sync

   # 或使用 pip
   pip install -e .
   ```
3. **配置文件**：

   ```bash
   # 复制配置模板
   cp config.example.yaml config.yaml

   # 编辑配置文件，填入 API Key
   ```
4. **启动服务**：

   ```bash
   # Web 服务模式（推荐）
   uv run python -m excel_agent.main serve

   # 命令行模式
   uv run python -m excel_agent.main cli --excel your_file.xlsx
   ```
5. **开始使用**：
   打开浏览器访问 `http://localhost:8000`

## ⚙️ 配置说明

### 配置文件模板

```yaml
model:
  # 当前使用的提供商
  provider: "openai"

  # 各提供商配置
  providers:
    openai:
      provider: "openai"
      model_name: "gpt-4"
      api_key: "${OPENAI_API_KEY}"
      base_url: "https://api.openai.com/v1"
      temperature: 0.1
      max_tokens: 4096
      fallback_models:
        - "gpt-3.5-turbo"

excel:
  max_preview_rows: 20
  default_result_limit: 20
  max_result_limit: 1000

server:
  host: "0.0.0.0"
  port: 8000
```

### 环境变量

也可使用环境变量配置：

```bash
export OPENAI_API_KEY="your-api-key"
export OPENAI_BASE_URL="your-api-base-url"
```

### 配置参数说明

| 参数                | 说明             | 示例                          |
| ------------------- | ---------------- | ----------------------------- |
| `provider`        | 使用的模型提供商 | `openai`                    |
| `model_name`      | 模型名称         | `gpt-4`                     |
| `api_key`         | API 密钥         | `sk-xxx`                    |
| `base_url`        | API 端点（可选） | `https://api.openai.com/v1` |
| `temperature`     | 温度参数         | `0.1`                       |
| `max_tokens`      | 最大 token 数    | `4096`                      |
| `fallback_models` | 降级模型列表     | `["gpt-3.5-turbo"]`         |

## 🎮 使用指南

### 1. 上传文件

- 拖拽 Excel 文件到页面
- 或点击上传区域选择文件
- 支持同时上传多个文件

### 2. 开始对话

在聊天框输入自然语言问题，例如：

- "这个表有多少行数据？"
- "按分公司统计销售总额"
- "帮我画个饼图展示各部门占比"
- "2024年11月的数据明细"

### 3. 使用示例

```
用户：这个表有多少行数据？
助手：该表共有 15,234 行数据。

用户：按分公司统计移动新增用户总数
助手：[调用 group_and_aggregate 工具]
      各分公司移动新增用户统计如下：
      | 分公司 | 移动新增用户 |
      |--------|-------------|
      | 东城   | 45,678      |
      | 西城   | 38,901      |
      | ...    | ...         |

用户：西城的明细呢？
助手：[理解上下文，调用 filter_data]
      西城分公司的详细数据如下：...

用户：用饼图展示各分公司的占比
助手：[调用 generate_chart 工具]
      📊 已生成饼图，共 8 个数据点。
      [交互式 ECharts 饼图显示]
```

## 📡 API 接口

启动服务后访问 `http://localhost:8000/docs` 查看完整 Swagger 文档。

### 主要接口

| 接口             | 方法 | 描述               |
| ---------------- | ---- | ------------------ |
| `/`            | GET  | Web 界面           |
| `/upload`      | POST | 上传 Excel 文件    |
| `/load`        | POST | 通过路径加载 Excel |
| `/chat/stream` | POST | 流式对话（推荐）   |
| `/chat`        | POST | 非流式对话         |
| `/status`      | GET  | 获取当前状态       |
| `/reset`       | POST | 重置 Agent         |

### 请求示例

```bash
# 上传 Excel
curl -X POST "http://localhost:8000/upload" \
  -F "file=@your_file.xlsx"

# 流式对话
curl -X POST "http://localhost:8000/chat/stream" \
  -H "Content-Type: application/json" \
  -d '{
    "message": "按部门统计销售额",
    "history": []
  }'
```

## 🔧 常见问题与解决方案

### Q1: 程序无法启动

**A: 环境问题**

- ✅ 确保 Python 版本 >= 3.11
- ✅ 确保已正确安装依赖：`uv sync` 或 `pip install -e .`
- ✅ 确保 `config.yaml` 文件存在且配置正确

### Q2: API 调用失败

**A: 配置问题**

- ✅ 检查 `api_key` 是否正确
- ✅ 检查 `base_url` 是否可访问
- ✅ 检查网络代理设置

### Q3: Excel 文件上传失败

**A: 文件问题**

- ✅ 确保文件格式为 `.xlsx` 或 `.xls`
- ✅ 检查文件是否损坏
- ✅ 检查文件大小是否过大

### Q4: 图表不显示

**A: 浏览器问题**

- ✅ 使用现代浏览器（Chrome、Firefox、Edge）
- ✅ 检查浏览器控制台是否有错误

### Q5: 多表联接失败

**A: 数据问题**

- ✅ 确保关联字段数据类型一致
- ✅ 检查字段名称是否正确

## 🐳 Docker 部署

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install uv && uv sync
EXPOSE 8000
CMD ["uv", "run", "python", "-m", "excel_agent.main", "serve"]
```

```bash
docker build -t excel-mind .
docker run -p 8000:8000 -e OPENAI_API_KEY=your-key excel-mind
```

## 🛠️ 开发说明

### 技术栈

- **Python 3.11** - 主要开发语言
- **LangGraph** - Agent 工作流框架
- **LangChain** - LLM 应用框架
- **FastAPI** - Web 服务框架
- **Pandas** - 数据处理
- **Chroma** - 向量数据库
- **ECharts** - 图表可视化

### 开发环境设置

```bash
# 安装开发依赖
uv sync --dev

# 运行测试
uv run pytest

# 代码格式化
uv run ruff format .
```

## 📝 更新日志

### v1.2.0 (当前版本)

- ✅ 新增 Anthropic 原生 API 支持
- ✅ 新增降级模型（fallback_models）配置
- ✅ 优化流式输出逻辑
- ✅ 更新依赖版本

### v1.1.0

- ✅ 迁移到 LangChain ReAct Agent
- ✅ 实现真正的流式输出
- ✅ 添加可折叠思考过程
- ✅ 修复工具描述和亮色模式 UI

### v1.0.0

- ✅ 基础 Excel 数据分析功能
- ✅ 自然语言查询
- ✅ 多轮对话支持
- ✅ ECharts 图表可视化
- ✅ 本地知识库

## 🤝 技术支持

如果遇到问题：

1. 检查配置文件是否正确
2. 查看控制台输出的错误信息
3. 提交 Issue 到 GitHub 仓库

## ☕ 赞赏支持

如果这个项目对您有帮助，欢迎请作者喝杯咖啡☕

**您的支持就是作者开发和维护项目的动力🚀**

### 微信赞赏码

![微信赞赏码](./docs/wechat_reward.jpg)

*扫码支持开发者，您的每一份心意都是我们前进的动力！*

## 📞 联系作者

### 💡 功能建议与合作需求

如果您有以下需求，欢迎联系作者：

- 🚀 **新功能建议** - 希望添加特定功能
- 🤝 **商业合作** - 定制开发需求
- 🛠️ **技术支持** - 专业问题解答
- 💼 **项目合作** - 相关项目合作机会

### 📱 联系方式

**微信号：** `wyh2353493891`

*添加微信时请备注：ExcelMind项目咨询*

## 📄 免责声明

- 本工具仅供学习和个人使用
- 请勿将敏感数据上传至公网部署的服务
- 作者不对因使用本工具导致的任何数据泄露负责
- AI 分析结果仅供参考，请自行验证重要数据

## 📄 License

[MIT License](LICENSE)

---

**📊 让数据分析变得简单！**

> 💡 提示：首次使用建议查看演示视频，快速了解各项功能。
