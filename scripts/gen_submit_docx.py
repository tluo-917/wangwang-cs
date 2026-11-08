# 生成项目技术说明 Word
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "docs" / "电商智能客服RAG系统技术文档.docx"

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
BLACK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x44, 0x44, 0x44)


def font(run, size=12, bold=False, name="Times New Roman", east="宋体", color=BLACK):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = name
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), east)


def add_p(doc, text, size=12, bold=False, align="left", space_after=8, first_line=True, east="宋体"):
    p = doc.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if first_line and align == "justify":
        pf.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    font(run, size=size, bold=bold, east=east)
    return p


def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(16 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(text)
    if level == 1:
        font(run, size=16, bold=True, east="黑体", color=NAVY)
    elif level == 2:
        font(run, size=14, bold=True, east="黑体", color=NAVY)
    else:
        font(run, size=12, bold=True, east="黑体", color=NAVY)
    return p


def body(doc, text):
    add_p(doc, text, align="justify", space_after=8)


def bullet(doc, text, level=0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.75 + level * 0.5)
    p.paragraph_format.first_line_indent = Cm(-0.37)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run("•  " + text)
    font(run, size=12)


def code_block(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    font(run, size=10, name="Consolas", east="宋体", color=GRAY)


def set_cell(cell, text, bold=False, center=False, size=10.5):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    run = p.add_run(text)
    font(run, size=size, bold=bold, east="宋体")


def shade_header(row):
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "1F3A5F")
        shd.set(qn("w:val"), "clear")
        tcPr.append(shd)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.bold = True


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        set_cell(table.rows[0].cells[i], h, bold=True, center=True, size=10.5)
        for run in table.rows[0].cells[i].paragraphs[0].runs:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shade_header(table.rows[0])
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            set_cell(table.rows[r].cells[c], val, center=(c == 0), size=10.5)
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return table


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.54)
    sec.right_margin = Cm(2.54)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # ---- 封面 ----
    for _ in range(3):
        add_p(doc, "", space_after=0, first_line=False)
    add_p(doc, "项目技术文档", size=14, align="center", space_after=24, first_line=False, east="黑体")
    add_p(doc, "基于大模型与 RAG 的电商智能客服系统", size=22, bold=True, align="center", space_after=8, first_line=False, east="黑体")
    add_p(doc, "技术说明文档", size=18, bold=True, align="center", space_after=28, first_line=False, east="黑体")
    add_p(doc, "（大模型接入、提示词工程与检索增强生成）", size=14, align="center", space_after=36, first_line=False, east="楷体")

    add_table(
        doc,
        ["项目", "内容"],
        [
            ["系统名称", "MewHelp（汪汪优选智能客服）"],
            ["文档类型", "技术说明"],
            ["覆盖范围", "大模型接入与提示词；知识库与检索增强"],
            ["技术栈", "FastAPI · LangChain · MySQL · Milvus · bge-m3 / bge-reranker"],
            ["日期", "2026 年 9 月"],
        ],
        col_widths=[4.2, 11.8],
    )

    doc.add_page_break()

    heading(doc, "1 引言", 1)
    heading(doc, "1.1 项目背景", 2)
    body(doc, "电商客服场景中，用户大量咨询集中在运费、退换货、时效和商品说明等政策类问题。仅依赖大模型参数记忆，容易出现过期政策、编造条款和无法追溯来源等问题。本系统在接入大模型完成流式对话与结构化提取的基础上，引入检索增强生成（Retrieval-Augmented Generation, RAG）：将平台政策与手册沉淀为可检索知识库，使回答建立在可引用证据之上；证据不足时明确拒答并引导人工，而不是生成无法核验的内容。")
    heading(doc, "1.2 文档范围", 2)
    body(doc, "这份说明写模型怎么接入、提示词怎么组织，以及知识库、混合检索和精排。订单查询、退款这些工具调用放在文末做边界说明。")
    heading(doc, "1.3 建设目标", 2)
    bullet(doc, "完成聊天模型、嵌入模型、重排模型三组上游的直连接入，配置与代码解耦。")
    bullet(doc, "采用对话式提示词模板，支持多轮上下文、角色约束与结构化输出。")
    bullet(doc, "建立 Markdown 知识库的切块、双写与向量化流程，保证原文可追溯。")
    bullet(doc, "实现 dense 向量检索与 BM25 关键词检索的多路融合，并经重排模型精排。")
    bullet(doc, "在对话中强制检索政策知识，对低置信结果拒答，回答可点选引用编号。")

    heading(doc, "2 总体架构", 1)
    heading(doc, "2.1 技术选型", 2)
    add_table(
        doc,
        ["层次", "技术选型", "说明"],
        [
            ["应用框架", "Python 3.12、FastAPI", "提供对话、提取与知识库页面接口"],
            ["模型接入", "LangChain ChatOpenAI", "兼容 OpenAI 协议，直连上游，无独立网关"],
            ["提示词", "ChatPromptTemplate", "按 system / human / history 组织多轮对话"],
            ["文本切分", "langchain-text-splitters", "Markdown 标题切分 + 中文递归切分"],
            ["嵌入模型", "BAAI/bge-m3", "向量维度 1024，用于稠密检索"],
            ["重排模型", "BAAI/bge-reranker-v2-m3", "对混合召回结果精排"],
            ["原文存储", "MySQL", "知识块权威源，记录切块正文与双写状态"],
            ["向量检索", "Milvus Standalone", "稠密向量 + 原生 BM25 稀疏检索 + hybrid_search"],
        ],
        col_widths=[3.2, 5.4, 7.4],
    )
    body(doc, "聊天、嵌入、重排三组上游分别配置基址、密钥与模型名（CHAT_*、EMBED_*、RERANK_*）。更换供应商时只修改环境变量，不修改业务代码。系统未使用 TextLoader、Chroma、VectorStore.as_retriever、MMR 以及 SelfQueryRetriever 等教程常用封装；检索入口为自研函数 search_knowledge，以便接入 Milvus 原生混合检索与独立精排服务。")

    heading(doc, "2.2 处理流程", 2)
    heading(doc, "2.2.1 对话接入流程", 3)
    body(doc, "用户在浏览器发送问题后，前端请求 POST /api/chat。服务端以 Server-Sent Events（SSE）将模型输出按 token 推送，事件形态为 data: {\"delta\": \"...\"}。提示词由 ChatPromptTemplate 渲染，包含客服系统角色与历史消息占位。售后描述可通过 POST /api/extract 抽取订单号、诉求类型与期望方案。")
    heading(doc, "2.2.2 知识入库流程", 3)
    body(doc, "离线或管理页将 data/kb 目录下的 Markdown 手册读入内存，按标题层级与长度约束切块，先写入 MySQL（状态 pending），再调用嵌入模型生成 1024 维向量，以相同主键 upsert 至 Milvus，回标状态 done。MySQL 保存全文与元数据，作为重建与对账的权威源；Milvus 作为在线检索读模型。")
    heading(doc, "2.2.3 在线 RAG 流程", 3)
    body(doc, "政策类问题在进入生成前执行强制检索：先将口语改写为标准问法并扩展同义词，再以 dense 与 BM25 各召回 Top-50，经 RRF 融合后由 reranker 精排出 Top-10。置信度或自评不足则返回拒答话术；通过则把编号证据注入上下文，模型作答时标注 [1]、[2]，前端将编号渲染为可点开来源。")

    heading(doc, "3 大模型接入与提示词工程", 1)
    heading(doc, "3.1 模型接入", 2)
    body(doc, "核心实现位于 app/core/llm.py。get_chat_model 基于 ChatOpenAI 创建客户端，传入模型名、base_url、api_key 与 streaming 开关。对话默认开启流式，便于客服界面逐字展示；结构化提取走 function calling，将结果绑定到 Pydantic 模型 AfterSalesTicket，避免正则或自由文本解析。嵌入调用独立的 OpenAI 兼容客户端，请求 BAAI/bge-m3；重排不走 Chat Completions，而是按 Jina / Cohere 形态请求 /rerank 接口。")
    heading(doc, "3.2 提示词组织", 2)
    body(doc, "全部模板集中在 app/core/prompts.py，统一使用 ChatPromptTemplate，而不使用面向单字符串填空的 PromptTemplate。原因在于客服任务是多轮对话：需要稳定的 system 角色、可变的历史消息以及当前用户输入。主要模板如下。")
    add_table(
        doc,
        ["模板名称", "形态", "用途"],
        [
            ["CUSTOMER_SERVICE_PROMPT", "system + MessagesPlaceholder(history)", "第一章纯对话客服人设与行为约束"],
            ["EXTRACT_PROMPT", "system + human{text}", "从售后描述抽取工单字段"],
            ["QUERY_REWRITE_PROMPT", "system + human{query}", "口语改写为标准问法，并给出 BM25 同义词"],
            ["SELF_CHECK_PROMPT", "system + human（问题与证据）", "生成前判断检索证据是否足够"],
            ["RAG_ANSWER_PROMPT", "system + human（问题与编号证据）", "约束模型仅依据证据作答并标注出处"],
            ["AGENT_SYSTEM", "多轮 system + history", "线上将检索证据注入对话后的作答原则"],
        ],
        col_widths=[4.6, 5.6, 5.8],
    )
    body(doc, "Query 改写失败时回退为用户原话继续检索，避免因提示词链路异常导致整轮请求失败。同义词扩展只拼接到 BM25 检索文本，不改写 dense 查询，以免把向量语义带偏。线上知识问答通常将编号证据作为本轮上下文注入，而不是每次单独调用 RAG_ANSWER_PROMPT；该模板更多用于生成约束与评估口径对齐。")

    heading(doc, "4 知识库构建", 1)
    heading(doc, "4.1 文档加载", 2)
    body(doc, "本系统未使用 LangChain 文档加载器。知识源为项目内 Markdown 文件，清单由 app/kb/sources.py 统一维护，离线脚本 scripts/build_kb.py 与知识库页面共用该清单。读取方式为 pathlib.Path.read_text。采用该方式的原因是语料格式固定、编码可控，引入通用 Loader 并不能增加信息，反而增加路径与解析差异。")
    heading(doc, "4.2 切块策略", 2)
    body(doc, "切块编排在 app/kb/documents.py 的 build_chunks，默认 chunk_size=400、overlap=60、表格最大行数 10。处理顺序为：")
    bullet(doc, "使用 MarkdownHeaderTextSplitter 按一级至四级标题切节，生成 section_path。")
    bullet(doc, "若当前节为 Markdown 表格，则按行切分，每块复制表头，避免半张表不可读。")
    bullet(doc, "否则使用 RecursiveCharacterTextSplitter 递归切分。中文分隔符依次为段落、换行、句末标点、逗号、空格，最后逐字。递归阶段 overlap 设为 0。")
    bullet(doc, "再由 apply_sentence_overlap 从上一块末尾截取完整句子作为重叠，目标长度不超过 60 字，避免半截句进入向量。")
    body(doc, "政策与手册类文本往往没有自然语言问句，此时 questions 字段落在章节标题，category 落在上级路径。关键条款通过标题与正文前部是否命中退款、运费、保修等词表进行标记，供后续置信度计算使用。")
    heading(doc, "4.3 双写与向量化", 2)
    body(doc, "app/kb/dualwrite.py 先将切块写入 MySQL 并串联前后块指针，状态为 pending；随后批量拼接 category、questions、answer 调用嵌入模型，按主键写入 Milvus 的 dense 向量与 BM25 文本字段，回填 vector_id 并将状态改为 done。中断后可只对 pending 块重跑向量化。Milvus 集合字段包括 id、dense（1024 维）、text / sparse（BM25）、question、answer、section_path、content_type、category。稠密索引为 AUTOINDEX + COSINE，稀疏索引为 SPARSE_INVERTED_INDEX + BM25。")

    heading(doc, "5 检索与生成控制", 1)
    heading(doc, "5.1 多路融合检索", 2)
    body(doc, "检索编排位于 app/core/retrieval.py 的 search_knowledge，支持四种策略，评估与线上共用同一实现：vector（仅稠密）、bm25（仅关键词）、hybrid（两路召回后 RRF）、hybrid_rerank（RRF 后再精排）。线上默认 hybrid_rerank。")
    body(doc, "混合检索在 app/kb/milvus_client.py 中调用 Milvus hybrid_search：dense 与 BM25 各召回 Top-50，使用 RRFRanker 融合。RRF（Reciprocal Rank Fusion）按公式 score = Σ 1/(k + rank) 合并排名，不直接相加 COSINE 与 BM25 的原始分，从而避免量纲不一致。本系统 RRFRanker 未设置额外权重，两路等权。融合结果再交给 bge-reranker-v2-m3 精排出 Top-10。")
    body(doc, "同义词只作用于 BM25 侧。对于一问多意图的长问句，系统按标点拆成子句分别检索，再按块编号轮转合并，避免单一意图占满名额。精排后将最相关结果置于上下文首部、次相关置于尾部，缓解大模型对中间段落不敏感的问题。")
    heading(doc, "5.2 强制检索与拒答", 2)
    body(doc, "知识类问题在图编排节点 retrieve_knowledge 中强制走检索，不把“是否查库”交给模型自由决定。检索后计算证据置信度，并可用自评模板判断证据是否足以回答。低于阈值则返回固定拒答，并提示联系人工；通过则生成带编号的 citations，前端将答案中的 [n] 渲染为可点击来源浮层。该机制用于降低政策幻觉。")

    heading(doc, "6 模块怎么拆", 1)
    body(doc, "整条链路按数据往下走：文档先变成可检索的知识块，再排成命中列表，最后由模型和提示词写成带引用的回答。聊天页只负责把接口里的事件画出来。")

    heading(doc, "6.1 模块对照", 2)
    add_table(
        doc,
        ["模块", "做什么", "主要源码", "怎么看效果"],
        [
            ["知识库与聊天页", "文档加载、切块、双写；聊天页 UI", "app/kb/chunking.py\napp/kb/documents.py\napp/kb/dualwrite.py\napp/static/index.html", "切块可预览、双写条数一致；聊天页流式出字、引用角标能点开"],
            ["检索", "四策略检索、RRF 融合、精排、评估", "app/core/retrieval.py\napp/kb/milvus_client.py\napp/core/embeddings.py\napp/core/rerank.py", "同一查询可切换 vector、bm25、hybrid、hybrid_rerank"],
            ["模型与提示词", "模型客户端、提示词、强制检索与对话接口", "app/core/llm.py\napp/core/prompts.py\napp/graph/nodes.py\napp/api/chat.py", "政策问答带编号引用；知识库外问题拒答"],
        ],
        col_widths=[2.4, 3.4, 4.6, 5.6],
    )

    heading(doc, "6.2 知识库与聊天页", 2)
    body(doc, "入库侧维护知识源清单、切块，并把同一批块写入 MySQL 和 Milvus。聊天页在 app/static/index.html：消息气泡、读取 POST /api/chat 的 SSE 流并逐字渲染 delta、把回答里的 [n] 做成可点击引用，以及满意度按钮。页面本身不做检索和生成，只消费 delta、citations、done。")
    heading(doc, "6.3 混合检索与精排", 2)
    body(doc, "在线检索支持四种策略，Milvus 侧配置稠密和稀疏索引，hybrid_search 之后接精排。评估和线上共用 search_knowledge。知识库页可以自测检索，评估页对比各策略的 Recall 和 MRR。")
    heading(doc, "6.4 模型接入与提示词", 2)
    body(doc, "模型客户端按场景分了三组上游。提示词集中在 ChatPromptTemplate，包括问题改写和证据是否够用的判断。对话接口用 SSE 往外推事件。知识类问题先强制检索，证据不够就拒答，不让模型凭空编政策。")

    heading(doc, "7 关键实现对照", 1)
    add_table(
        doc,
        ["问题", "结论", "代码位置"],
        [
            ["使用何种加载器", "不使用 LangChain Loader，直接读取 Markdown", "scripts/build_kb.py"],
            ["分块大小与重叠", "chunk_size=400，句级 overlap=60", "app/kb/documents.py"],
            ["是否多路融合", "是。dense + BM25，RRF 融合后再精排", "app/kb/milvus_client.py\napp/core/retrieval.py"],
            ["提示词模板类型", "全部为 ChatPromptTemplate", "app/core/prompts.py"],
            ["向量维度", "1024（bge-m3）", "app/kb/milvus_client.py"],
            ["线上默认策略", "hybrid_rerank", "search_knowledge 默认参数"],
            ["流式如何展示", "后端 SSE 推 delta/citations；聊天页逐字渲染并可点引用", "app/api/chat.py\napp/static/index.html"],
        ],
        col_widths=[3.6, 6.6, 5.8],
    )

    heading(doc, "8 验证与演示要点", 1)
    bullet(doc, "对话页验证流式输出与多轮上下文保持。")
    bullet(doc, "提取接口验证售后描述可得到订单号、诉求类型、期望方案。")
    bullet(doc, "知识库页验证切块预览、双写一致，并用“邮费是多少”对比四路检索。")
    bullet(doc, "对话页用政策问题验证编号引用；用知识库外问题验证拒答。")
    bullet(doc, "评估页展示四策略对比结果，说明评测调用的即线上检索函数。")

    heading(doc, "9 总结", 1)
    body(doc, "本系统将大模型接入与 RAG 管线结合，形成可演示、可追溯的电商客服知识问答能力。接入层以直连方式和 ChatPromptTemplate 完成流式对话与结构化提取；知识层以可控的 Markdown 切块和双写保证原文权威性；检索层以稠密向量与关键词多路融合并经精排提高命中质量；生成层以强制检索和拒答降低幻觉；交互层由聊天页消费 SSE 并展示引用。入库、检索、提示词和聊天页按这条链路接在一起，回答可以回溯到原文。")

    heading(doc, "附录 范围说明", 1)
    body(doc, "系统里还有订单、物流这类工具调用。这份说明主要写模型接入、提示词和检索增强。政策类问题的主路径是先查知识库，而不是让模型自己决定要不要查。")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
