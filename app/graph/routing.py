import re

from langchain_core.messages import AIMessage, HumanMessage

from app.config import settings

# 九类意图 → 五出口(写死的分流规则,spec §3.1 / README route_by_intent;
# 返回值 = build.py 条件边映射键,单一来源不漂移)
INTENT_TO_ROUTE: dict[str, str] = {
    "投诉": "escalate",
    "闲聊": "fallback_script",
    "其他": "fallback_script",
    "商品咨询": "knowledge",
    "退款退货": "refund_flow",
    "售后": "refund_flow",
    "人工": "business",   # 明确要求建工单/转人工 → 主力 Agent 走 create_ticket 确认流
    "物流": "business",
    "订单": "business",
}


# 价格/库存是 query_product 的实时数据,知识库里没有;「有货」要配「吗/还/没有」以免误伤政策句。
_LIVE_PRODUCT_RE = re.compile(
    r"(多少钱|什么价|售价|单价|报价|有货吗|还有货|有没有货|没货了|库存|现货|缺货|还能买)"
)


def is_live_product_query(text: str | None) -> bool:
    """问的是实时价格或库存,不是政策 FAQ。"""
    return bool(text and _LIVE_PRODUCT_RE.search(text))


def _query_text(state) -> str:
    q = (state.get("resolved_query") or "").strip()
    if q:
        return q
    for m in reversed(state.get("messages") or []):
        if isinstance(m, HumanMessage) and isinstance(m.content, str) and m.content.strip():
            return m.content
    return ""


def route_by_intent(state) -> str:
    """按意图分流;未知意图保守归 business。

 「商品咨询」默认走知识库强制检索,但价格/库存不在库里——再走 RAG 只会误拒答,
 下一轮问订单时 Agent 还会把那句没答完的商品问题一起补上。这类问句改走 business,
 让 query_product 接手。政策类商品咨询(运费、退货规则)仍走 knowledge。
 """
    intent = state.get("intent", "")
    if intent == "商品咨询" and is_live_product_query(_query_text(state)):
        return "business"
    return INTENT_TO_ROUTE.get(intent, "business")


def confidence_gate(state) -> str:
    """知识路生成前证据闸:强放行、弱兜底。"""
    return "strong" if state.get("evidence_strong") else "weak"


def should_continue(state) -> str:
    """ReAct 停止条件:无 tool_calls 就收敛,步数封顶则强停,否则继续。

 只用步数封顶,不在这里卡 token。防打转靠的是步数:一步一次模型调用,数得清、好解释,
 各家 Agent SDK 给的也是 max_turns / max_iterations 这类步数上限。
 token 花销是另一件事——它管的是烧钱和撑爆上下文,不是死循环,阈值要按真实用量分布标定,
 归第 9 章的成本控制一起讲。混在这里卡,一次估偏就把工具链整条掐断,而且掐在模型
 已经决定调工具之后,用户只看到半句「我这就去查」。"""
    last = state["messages"][-1]
    has_tool_calls = isinstance(last, AIMessage) and bool(last.tool_calls)
    if not has_tool_calls:
        return "stop"
    if state.get("steps", 0) >= settings.max_agent_steps:
        return "stop"
    return "continue"
