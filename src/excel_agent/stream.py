"""流式对话 - 使用 LangChain ReAct Agent

v2.0: 集成 SkillManager 动态工具路由
"""

import json
from typing import Any, AsyncGenerator, Dict, List, Optional

from langchain_core.messages import HumanMessage, SystemMessage, AIMessageChunk, ToolMessage
from langchain_openai import ChatOpenAI
from langchain_anthropic import ChatAnthropic
from langgraph.prebuilt import create_react_agent

from .config import get_config
from .excel_loader import get_loader
from .knowledge_base import get_knowledge_base, format_knowledge_context
from .tools import ALL_TOOLS
from .skills import get_skill_manager


# 当前使用的模型名称（用于容错切换后的显示）
_current_model: Optional[str] = None


class CustomJSONEncoder(json.JSONEncoder):
    """自定义 JSON 编码器，处理 Pandas/Numpy 类型"""

    def default(self, obj):
        # 处理 Pandas Timestamp
        if hasattr(obj, 'isoformat'):
            return obj.isoformat()
        # 处理 numpy 类型
        if hasattr(obj, 'item'):
            return obj.item()
        # 处理 numpy 数组
        if hasattr(obj, 'tolist'):
            return obj.tolist()
        # 处理 pandas NaT
        if str(obj) == 'NaT':
            return None
        # 处理 pandas NA
        if str(obj) == '<NA>':
            return None
        return super().default(obj)


def json_dumps(obj, **kwargs):
    """使用自定义编码器的 JSON 序列化函数"""
    return json.dumps(obj, cls=CustomJSONEncoder, **kwargs)


SYSTEM_PROMPT = """你是一个专业的 Excel 数据分析助手。

## 当前 Excel 信息
{excel_summary}

## 相关知识参考
{knowledge_context}

## 当前可用技能
{skills_context}

## 工作原则
1. 根据用户问题，判断是否需要使用工具
2. 如需工具，调用合适的工具获取数据
3. 工具调用成功后，根据结果回答用户问题
4. **最终回答直接给出结论和分析**，不要描述"我使用了xx工具"或"我进行了xx操作"等内部过程
5. 回答语气友好，使用中文，并给出自己的一些数据分析建议
6. 如果有相关知识参考，请遵循其中的规则和建议

## 重要：完成后必须总结
当你完成用户请求的所有操作后，**必须**给出简洁的完成总结，包括：
- 已完成的操作概述
- 关键数据结果（如统计汇总值）
- 提示用户可以点击页面左侧的"下载文件"按钮获取修改后的文件
"""


# 是否启用 Skills 动态路由（可通过配置控制）
# 设为 False 让 LLM 自己判断使用哪些工具，而不是通过关键词预筛选
ENABLE_SKILL_ROUTING = False


def create_llm_for_model(model_name: str):
    """根据模型名称创建 LLM 实例"""
    config = get_config()
    provider = config.model.get_active_provider()

    # 根据模型名称自动选择 provider 类型
    # gemini/gpt 模型使用 OpenAI 兼容接口，claude 模型使用 Anthropic 接口
    use_openai = model_name.startswith(("gemini", "gpt", "deepseek", "qwen", "glm"))

    # 处理 base_url
    base_url = provider.base_url or ""

    if use_openai:
        # OpenAI 格式需要确保有 /v1 后缀
        openai_base_url = base_url.rstrip('/')
        if not openai_base_url.endswith('/v1'):
            openai_base_url = openai_base_url + '/v1'
        return ChatOpenAI(
            model=model_name,
            api_key=provider.api_key,
            base_url=openai_base_url,
            temperature=provider.temperature,
            max_tokens=provider.max_tokens,
        )
    elif provider.provider == "anthropic" or model_name.startswith("claude"):
        # Anthropic 格式需要去掉末尾的斜杠
        return ChatAnthropic(
            model=model_name,
            api_key=provider.api_key,
            base_url=base_url.rstrip('/') if base_url else None,
            temperature=provider.temperature,
            max_tokens=provider.max_tokens,
        )
    else:
        # 默认使用 OpenAI 兼容接口
        return ChatOpenAI(
            model=model_name,
            api_key=provider.api_key,
            base_url=base_url if base_url else None,
            temperature=provider.temperature,
            max_tokens=provider.max_tokens,
        )


def get_model_list() -> List[str]:
    """获取按优先级排序的模型列表（主模型 + 降级模型）"""
    config = get_config()
    provider = config.model.get_active_provider()

    models = [provider.model_name]
    if provider.fallback_models:
        models.extend(provider.fallback_models)

    return models


def get_llm():
    """获取 LLM 实例（使用主模型）"""
    config = get_config()
    provider = config.model.get_active_provider()
    return create_llm_for_model(provider.model_name)


async def stream_chat(message: str, history: list = None) -> AsyncGenerator[Dict[str, Any], None]:
    """执行对话 - 使用 LangChain ReAct Agent（支持模型容错）

    Args:
        message: 当前用户消息
        history: 历史对话列表，每项为 {"role": "user"|"assistant", "content": "..."}

    Yields:
        状态事件类型:
        - {"type": "status", "status": "processing"} - 开始处理，前端应禁用输入
        - {"type": "status", "status": "idle"} - 处理完成，前端可恢复输入
        - {"type": "thinking", "content": "..."} - 思考过程
        - {"type": "thinking_done"} - 思考完成
        - {"type": "tool_call", "name": "...", "args": {...}} - 工具调用
        - {"type": "tool_result", "name": "...", "result": {...}} - 工具结果
        - {"type": "token", "content": "..."} - 流式输出 token
        - {"type": "done", "content": "..."} - 回答完成
        - {"type": "error", "content": "..."} - 错误信息
    """
    global _current_model
    loader = get_loader()

    # 开始处理 - 通知前端禁用输入
    yield {"type": "status", "status": "processing"}

    if not loader.is_loaded:
        yield {"type": "error", "content": "请先上传 Excel 文件"}
        yield {"type": "status", "status": "idle"}
        return

    # 获取模型列表（主模型 + 降级模型）
    model_list = get_model_list()
    last_error = None

    for model_idx, model_name in enumerate(model_list):
        try:
            _current_model = model_name

            # 如果不是第一个模型，说明是降级
            if model_idx > 0:
                print(f"[模型容错] 切换到降级模型: {model_name}")
                yield {"type": "thinking", "content": f"模型负载过高，切换到 {model_name}..."}

            # 执行对话
            async for event in _do_stream_chat(message, history, model_name):
                yield event

            # 成功完成，发送 idle 状态后退出
            yield {"type": "status", "status": "idle"}
            return

        except Exception as e:
            error_str = str(e)
            last_error = e

            # 检查是否是可重试的错误（模型负载、连接问题、空响应等）
            is_retryable_error = (
                "500" in error_str or
                "负载" in error_str or
                "overload" in error_str.lower() or
                "capacity" in error_str.lower() or
                "rate" in error_str.lower() or
                "no generations" in error_str.lower() or  # Gemini 空流响应
                "empty" in error_str.lower() or
                "timeout" in error_str.lower() or
                "connection" in error_str.lower()
            )

            if is_retryable_error and model_idx < len(model_list) - 1:
                # 还有降级模型可用，继续尝试
                print(f"[模型容错] {model_name} 不可用: {error_str[:100]}")
                continue
            else:
                # 没有更多模型或不是负载错误，抛出异常
                import traceback
                traceback.print_exc()
                yield {"type": "thinking_done"}
                yield {"type": "error", "content": f"处理出错: {str(e)}"}
                yield {"type": "status", "status": "idle"}
                return

    # 所有模型都失败
    yield {"type": "thinking_done"}
    yield {"type": "error", "content": f"所有模型均不可用，最后错误: {str(last_error)}"}
    yield {"type": "status", "status": "idle"}


async def _do_stream_chat(message: str, history: list, model_name: str) -> AsyncGenerator[Dict[str, Any], None]:
    """执行单个模型的对话流程"""
    loader = get_loader()

    excel_summary = loader.get_summary()
    llm = create_llm_for_model(model_name)

    # 主对话开始
    yield {"type": "thinking", "content": f"正在使用 {model_name} 规划解答..."}

    # 检索相关知识
    knowledge_context = "暂无相关知识参考。"
    kb = get_knowledge_base()
    if kb:
        try:
            stats = kb.get_stats()
            print(f"[知识库] 状态: {stats['total_entries']} 条知识")
            relevant_knowledge = kb.search(query=message)
            print(f"[知识库] 检索到 {len(relevant_knowledge)} 条相关知识")
            if relevant_knowledge:
                knowledge_context = format_knowledge_context(relevant_knowledge)
                yield {"type": "thinking", "content": f"找到 {len(relevant_knowledge)} 条相关知识参考..."}
        except Exception as e:
            print(f"[知识库检索] 警告: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("[知识库] 未启用或初始化失败")

    # Skills 动态路由
    skills_context = "所有数据分析工具均可用。"
    active_tools = ALL_TOOLS  # 默认使用所有工具

    if ENABLE_SKILL_ROUTING:
        try:
            skill_manager = get_skill_manager()
            resolved_skills = skill_manager.resolve(message, top_k=5, threshold=0.25)

            if resolved_skills:
                active_tools = skill_manager.get_active_tools()
                skill_names = [s.display_name for s in resolved_skills]
                skills_context = f"已激活技能: {', '.join(skill_names)}\n"
                skills_context += skill_manager.get_system_prompt_additions()

                yield {"type": "thinking", "content": f"激活技能: {', '.join(skill_names)}"}
                print(f"[Skills] 激活: {skill_names}, 工具数: {len(active_tools)}")
        except Exception as e:
            print(f"[Skills] 路由失败，使用全部工具: {e}")
            active_tools = ALL_TOOLS

    # 构建系统提示
    system_prompt = SYSTEM_PROMPT.format(
        excel_summary=excel_summary,
        knowledge_context=knowledge_context,
        skills_context=skills_context
    )

    # 获取当前活跃表信息
    active_table_info = loader.get_active_table_info()
    current_table_name = active_table_info.filename if active_table_info else "未知表"

    # 创建 ReAct Agent（使用动态激活的工具）
    agent = create_react_agent(llm, active_tools)

    # 构建消息 - 包含历史对话
    current_message = f"[当前操作表: {current_table_name}] {message}"
    messages = [SystemMessage(content=system_prompt)]

    # 添加历史对话
    if history:
        from langchain_core.messages import AIMessage
        for msg in history:
            if msg.get("role") == "user":
                messages.append(HumanMessage(content=msg.get("content", "")))
            elif msg.get("role") == "assistant":
                messages.append(AIMessage(content=msg.get("content", "")))

    # 添加当前用户消息
    messages.append(HumanMessage(content=current_message))

    # 使用 stream_mode="messages" 获取真正的流式输出
    thinking_content = ""
    final_content = ""
    tool_call_yielded = False
    thinking_done_sent = False

    # 累积工具调用信息
    tool_names_by_id = {}  # id -> name
    args_by_index = {}  # index -> args_str
    tool_call_order = []  # 记录工具调用的顺序 [(id, index), ...]
    yielded_tool_ids = set()

    async for chunk in agent.astream(
        {"messages": messages},
        stream_mode="messages",
        config={"recursion_limit": 50}
    ):
        # chunk 是一个 tuple: (message, metadata)
        if isinstance(chunk, tuple) and len(chunk) >= 2:
            msg, metadata = chunk[0], chunk[1]

            # 处理 AIMessageChunk (LLM 输出)
            if isinstance(msg, AIMessageChunk):
                content = msg.content if hasattr(msg, 'content') else ""
                # 处理 content 可能是列表的情况（Anthropic 格式）
                if isinstance(content, list):
                    text_parts = []
                    for block in content:
                        if isinstance(block, dict) and block.get("type") == "text":
                            text_parts.append(block.get("text", ""))
                        elif isinstance(block, str):
                            text_parts.append(block)
                    content = "".join(text_parts)
                tool_call_chunks = getattr(msg, 'tool_call_chunks', [])

                # 累积工具调用的 chunks
                if tool_call_chunks:
                    for tcc in tool_call_chunks:
                        tc_id = tcc.get("id")
                        tc_name = tcc.get("name", "")
                        tc_args = tcc.get("args", "")
                        tc_index = tcc.get("index", 0)

                        # 如果有新的工具 id 出现（带 name），记录下来
                        if tc_id and tc_name:
                            if tc_id not in tool_names_by_id:
                                tool_call_order.append((tc_id, tc_index))
                                if not thinking_done_sent:
                                    yield {"type": "thinking_done"}
                                    thinking_done_sent = True
                            tool_names_by_id[tc_id] = tc_name

                        # 累积 args（按 index）
                        if tc_args:
                            if tc_index not in args_by_index:
                                args_by_index[tc_index] = ""
                            args_by_index[tc_index] += tc_args

                # 如果有文本内容
                if content:
                    if tool_call_yielded:
                        # 工具调用后的内容是最终回答
                        final_content += content
                        yield {"type": "token", "content": content}
                    else:
                        # 工具调用前的内容是思考过程
                        thinking_content += content
                        yield {"type": "thinking", "content": thinking_content}

            # 处理 ToolMessage (工具结果)
            elif isinstance(msg, ToolMessage):
                tool_call_id = msg.tool_call_id if hasattr(msg, 'tool_call_id') else None
                tool_name = msg.name if hasattr(msg, 'name') else "tool"
                tool_content = msg.content

                # 在发送 tool_result 之前，先发送对应的 tool_call
                if tool_call_id and tool_call_id not in yielded_tool_ids:
                    yielded_tool_ids.add(tool_call_id)
                    tool_call_yielded = True

                    # 找到这个工具调用的 index
                    tc_index = 0
                    for (tid, idx) in tool_call_order:
                        if tid == tool_call_id:
                            tc_index = idx
                            break

                    # 从 args_by_index 获取 args
                    args_str = args_by_index.get(tc_index, "{}")

                    # 解析 args
                    try:
                        args = json.loads(args_str)
                    except json.JSONDecodeError:
                        try:
                            last_brace = args_str.rfind('{"')
                            if last_brace >= 0:
                                args = json.loads(args_str[last_brace:])
                            else:
                                args = {"raw": args_str}
                        except:
                            args = {"raw": args_str}

                    # 获取工具名称
                    tc_name = tool_names_by_id.get(tool_call_id, tool_name)

                    yield {
                        "type": "tool_call",
                        "name": tc_name,
                        "args": args,
                    }

                    # 清除已处理的 args
                    if tc_index in args_by_index:
                        del args_by_index[tc_index]

                # 发送工具结果
                try:
                    result = json.loads(tool_content)
                except:
                    result = {"result": tool_content}

                yield {
                    "type": "tool_result",
                    "name": tool_name,
                    "result": result,
                }

    # 流式结束
    if not tool_call_yielded and thinking_content:
        # 没有工具调用，thinking_content 就是最终回答
        if not thinking_done_sent:
            yield {"type": "thinking_done"}
        yield {"type": "clear_thinking"}
        yield {"type": "token", "content": thinking_content}
        yield {"type": "done", "content": thinking_content}
    elif final_content:
        yield {"type": "done", "content": final_content}
    elif tool_call_yielded:
        # 工具调用后没有生成总结，添加后备提示
        fallback_message = "\n\n✅ **操作已完成**\n\n如需获取修改后的文件，请点击页面左侧的「下载文件」按钮。"
        yield {"type": "token", "content": fallback_message}
        yield {"type": "done", "content": fallback_message}
    else:
        yield {"type": "done", "content": ""}
