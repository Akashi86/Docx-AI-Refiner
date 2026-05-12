import concurrent.futures
import csv
import difflib
import html
import io
import json
import re
import time
import zipfile
import xml.etree.ElementTree as ET

import requests
import streamlit as st


NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

W_NS = f"{{{NAMESPACES['w']}}}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
SEGMENT_RE = re.compile(r"<seg\s+id=\"(\d+)\">(.*?)</seg>", re.DOTALL)


st.set_page_config(page_title="AI Word 论文润色工具", layout="wide")

st.markdown(
    """
    <style>
    .log-box {
        background-color: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 8px;
        font-family: Consolas, monospace;
        font-size: 13px;
        height: 350px;
        overflow-y: auto;
        border: 1px solid #333;
    }
    .log-entry { margin-bottom: 5px; word-wrap: break-word; }
    .log-success { color: #4ade80; font-weight: bold; }
    .log-send { color: #60a5fa; }
    .log-warn { color: #fbbf24; }
    .log-err { color: #f87171; font-weight: bold; }
    .log-info { color: #c084fc; font-weight: bold; }
    .diff-box {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 12px;
        background: #fafafa;
        line-height: 1.7;
    }
    .diff-del {
        color: #b91c1c;
        background: #fee2e2;
        text-decoration: line-through;
        padding: 1px 3px;
        border-radius: 4px;
    }
    .diff-ins {
        color: #166534;
        background: #dcfce7;
        text-decoration: none;
        padding: 1px 3px;
        border-radius: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "logs" not in st.session_state:
    st.session_state.logs = []
if "processed_file" not in st.session_state:
    st.session_state.processed_file = None
if "rewrite_report" not in st.session_state:
    st.session_state.rewrite_report = []


def add_log(message, kind="normal"):
    safe_message = html.escape(str(message))
    time_str = time.strftime("%H:%M:%S")
    class_map = {
        "success": "log-success",
        "send": "log-send",
        "warn": "log-warn",
        "err": "log-err",
        "info": "log-info",
    }
    css_class = class_map.get(kind, "")
    class_attr = f" {css_class}" if css_class else ""
    st.session_state.logs.append(
        f'<div class="log-entry{class_attr}">[{time_str}] {safe_message}</div>'
    )


def render_logs(log_container):
    log_container.markdown(
        f'<div class="log-box">{"".join(st.session_state.logs)}</div>',
        unsafe_allow_html=True,
    )


def build_segment_payload(text_nodes):
    parts = []
    for idx, elem in enumerate(text_nodes):
        text = elem.text or ""
        parts.append(f'<seg id="{idx}">{html.escape(text, quote=False)}</seg>')
    return "".join(parts)


def parse_segment_response(response_text, expected_count):
    matches = SEGMENT_RE.findall(response_text.strip())
    if len(matches) != expected_count:
        raise ValueError(f"AI 返回了 {len(matches)} 个片段，但期望 {expected_count} 个。")

    values = [None] * expected_count
    seen = set()
    for raw_id, raw_value in matches:
        idx = int(raw_id)
        if idx < 0 or idx >= expected_count:
            raise ValueError(f"AI 返回了非法片段 id: {idx}")
        if idx in seen:
            raise ValueError(f"AI 重复返回了片段 id: {idx}")
        seen.add(idx)
        values[idx] = html.unescape(raw_value)

    missing = [idx for idx, value in enumerate(values) if value is None]
    if missing:
        raise ValueError(f"AI 缺少片段 id: {missing}")
    return values


def make_diff_html(old_text, new_text):
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    parts = []

    for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        old_part = " ".join(old_words[old_start:old_end])
        new_part = " ".join(new_words[new_start:new_end])

        if tag == "equal":
            parts.append(html.escape(new_part))
        elif tag == "delete":
            parts.append(f'<del class="diff-del">{html.escape(old_part)}</del>')
        elif tag == "insert":
            parts.append(f'<ins class="diff-ins">{html.escape(new_part)}</ins>')
        elif tag == "replace":
            parts.append(f'<del class="diff-del">{html.escape(old_part)}</del>')
            parts.append(f'<ins class="diff-ins">{html.escape(new_part)}</ins>')

    return " ".join(part for part in parts if part)


def report_to_json(report):
    return json.dumps(report, ensure_ascii=False, indent=2).encode("utf-8")


def report_to_csv(report):
    output = io.StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=["paragraph_index", "page", "original_text", "new_text", "status", "error"],
    )
    writer.writeheader()
    for item in report:
        writer.writerow(
            {
                "paragraph_index": item.get("paragraph_index", ""),
                "page": item.get("page", ""),
                "original_text": item.get("original_text", ""),
                "new_text": item.get("new_text", ""),
                "status": item.get("status", ""),
                "error": item.get("error", ""),
            }
        )
    return output.getvalue().encode("utf-8-sig")


def make_system_prompt(user_prompt):
    return f"""{user_prompt}

你会收到一段来自 Word 文档的 XML 片段化文本，格式如下：
<seg id="0">...</seg><seg id="1">...</seg>

请严格遵守：
1. 只润色每个 <seg> 标签内部的正文。
2. 必须保留所有 <seg id="..."> 和 </seg> 标签。
3. 必须保持片段数量、id 数字和顺序完全不变。
4. 不要添加解释、Markdown、代码块或额外文本。
5. 不要合并、删除、拆分或重排任何 seg 标签。
6. 标签外不能输出任何内容。
"""


def call_deepseek(segment_payload, user_prompt, api_key, model_name, task_id, max_retries=3):
    url = "https://api.deepseek.com/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": make_system_prompt(user_prompt)},
            {"role": "user", "content": segment_payload},
        ],
        "temperature": 0.2,
    }

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=120)
            if response.status_code == 401:
                raise RuntimeError("API Key 错误或无效。")
            if response.status_code == 402:
                raise RuntimeError("账号余额不足。")
            if response.status_code in (400, 404):
                raise RuntimeError(f"模型 {model_name} 不可用或请求格式错误。")
            if response.status_code == 429 or response.status_code >= 500:
                raise requests.RequestException(
                    f"DeepSeek 暂时不可用或限流，HTTP {response.status_code}: {response.text[:300]}"
                )
            response.raise_for_status()

            data = response.json()
            return data["choices"][0]["message"]["content"].strip()
        except RuntimeError:
            raise
        except Exception as exc:
            last_error = exc
            if attempt == max_retries:
                break
            time.sleep(min(attempt * 4, 20))

    raise RuntimeError(f"段落 {task_id} 调用失败，已重试 {max_retries} 次：{last_error}")


def paragraph_text_nodes(paragraph):
    nodes = []
    for text_elem in paragraph.iter(f"{W_NS}t"):
        if text_elem.text:
            nodes.append(text_elem)
    return nodes


def paragraph_plain_text(paragraph):
    return "".join(elem.text or "" for elem in paragraph.iter(f"{W_NS}t")).strip()


def paragraph_page_number(paragraph, current_page):
    page_breaks = len(paragraph.findall(f".//{W_NS}lastRenderedPageBreak"))
    for br in paragraph.findall(f".//{W_NS}br"):
        if br.get(f"{W_NS}type") == "page":
            page_breaks += 1
    return current_page + page_breaks


def is_body_paragraph(paragraph):
    text = paragraph_plain_text(paragraph)
    if not text:
        return False
    return True


def collect_tasks(root, start_page, end_page, min_chars):
    tasks = []
    current_page = 1
    end_page_num = end_page if end_page else float("inf")

    paragraphs = list(root.iter(f"{W_NS}p"))
    for idx, paragraph in enumerate(paragraphs):
        current_page = paragraph_page_number(paragraph, current_page)
        if current_page < start_page or current_page > end_page_num:
            continue
        if not is_body_paragraph(paragraph):
            continue

        text_nodes = paragraph_text_nodes(paragraph)
        plain_text = "".join(elem.text or "" for elem in text_nodes).strip()
        if len(plain_text) < min_chars:
            continue

        tasks.append(
            {
                "paragraph_index": idx,
                "page": current_page,
                "text_nodes": text_nodes,
                "segment_payload": build_segment_payload(text_nodes),
                "segment_count": len(text_nodes),
                "plain_text": plain_text,
            }
        )
    return tasks


def apply_segment_values(task, values):
    for text_elem, new_text in zip(task["text_nodes"], values):
        text_elem.text = new_text
        if new_text.startswith(" ") or new_text.endswith(" "):
            text_elem.set(XML_SPACE, "preserve")


def process_word(
    file_bytes,
    api_key,
    model_name,
    concurrency,
    start_page,
    end_page,
    min_chars,
    prompt,
    log_container,
    progress_bar,
):
    st.session_state.logs = []
    st.session_state.rewrite_report = []

    try:
        add_log("正在读取并解析 Word 文档。")
        render_logs(log_container)

        input_buffer = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_buffer, "r") as doc_zip:
            xml_content = doc_zip.read("word/document.xml")
            root = ET.fromstring(xml_content)

            add_log("正在扫描正文段落并生成 AI 任务。")
            tasks = collect_tasks(root, start_page, end_page, min_chars)
            if not tasks:
                add_log("在设定范围内没有找到符合长度要求的正文段落。", "warn")
                render_logs(log_container)
                return None

            add_log(f"共找到 {len(tasks)} 个待处理段落，开始 {concurrency} 并发处理。", "info")
            render_logs(log_container)

            completed_tasks = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=concurrency) as executor:
                future_to_task = {}
                for task in tasks:
                    future = executor.submit(
                        call_deepseek,
                        task["segment_payload"],
                        prompt,
                        api_key,
                        model_name,
                        task["paragraph_index"] + 1,
                    )
                    future_to_task[future] = task
                    add_log(
                        f"第 {task['paragraph_index'] + 1} 段，第 {task['page']} 页，已发送给模型。",
                        "send",
                    )
                    render_logs(log_container)

                for future in concurrent.futures.as_completed(future_to_task):
                    task = future_to_task[future]
                    para_no = task["paragraph_index"] + 1
                    try:
                        response_text = future.result()
                        values = parse_segment_response(response_text, task["segment_count"])
                        original_text = task["plain_text"]
                        new_text = "".join(values).strip()
                        apply_segment_values(task, values)
                        st.session_state.rewrite_report.append(
                            {
                                "paragraph_index": para_no,
                                "page": task["page"],
                                "original_text": original_text,
                                "new_text": new_text,
                                "diff_html": make_diff_html(original_text, new_text),
                                "status": "success",
                                "error": "",
                            }
                        )
                        add_log(f"第 {para_no} 段处理完成，并已写回原文本节点。", "success")
                    except Exception as exc:
                        st.session_state.rewrite_report.append(
                            {
                                "paragraph_index": para_no,
                                "page": task["page"],
                                "original_text": task["plain_text"],
                                "new_text": task["plain_text"],
                                "diff_html": "",
                                "status": "failed",
                                "error": str(exc),
                            }
                        )
                        add_log(f"第 {para_no} 段处理失败，已保留原文：{exc}", "err")

                    completed_tasks += 1
                    progress_bar.progress(completed_tasks / len(tasks))
                    render_logs(log_container)

            output_io = io.BytesIO()
            with zipfile.ZipFile(output_io, "w", zipfile.ZIP_DEFLATED) as out_zip:
                for item in doc_zip.infolist():
                    if item.filename != "word/document.xml":
                        out_zip.writestr(item, doc_zip.read(item.filename))

                xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                out_zip.writestr("word/document.xml", xml_str)

        add_log("新文档已打包完成。原段落、run 和样式结构没有被重建。", "success")
        render_logs(log_container)
        return output_io.getvalue()

    except Exception as exc:
        add_log(f"处理流程发生异常：{exc}", "err")
        render_logs(log_container)
        return None


st.title("AI Word 论文逐段润色工具")
st.markdown("上传 .docx 后，系统会逐段发送正文给 AI，并只替换 Word XML 里的文本节点，尽量保持原有格式结构不变。")

col_left, col_right = st.columns([1, 1.2])

with col_left:
    st.subheader("1. 模型与处理范围")
    api_key = st.text_input("DeepSeek API Key", type="password")

    col_1, col_2 = st.columns(2)
    model_name = col_1.selectbox("模型", ["deepseek-chat", "deepseek-reasoner"])
    concurrency = col_2.slider(
        "并发请求数",
        min_value=1,
        max_value=8,
        value=3,
        help="建议 2-4。并发过高可能触发 API 限流。",
    )

    col_3, col_4 = st.columns(2)
    start_page = col_3.number_input("开始页", min_value=1, value=1)
    end_page_input = col_4.number_input("结束页，0 表示到末尾", min_value=0, value=0)
    end_page = int(end_page_input) if end_page_input > 0 else None

    min_chars = st.slider("忽略短段落，少于 N 字符不处理", min_value=5, max_value=120, value=20, step=5)

    st.subheader("2. 润色提示词")
    default_prompt = """Without changing the paragraph structure or the order of information, rewrite the text below to reduce Al vibes. Do not start sentences with generic stock phrases ("in conclusion," "moreover," etc.); use more specific, context-tied transitions instead. Make each paragraph's first sentence feel like a natural continuation of the previous one. Favor concrete verbs over stacks of abstract nouns, keep terminology precise, and split long sentences into 2- 3 shorter ones."""
    prompt = st.text_area("提示词", value=default_prompt, height=180)

    st.subheader("3. 上传文档")
    uploaded_file = st.file_uploader("选择 Word 文档 (.docx)", type=["docx"])

with col_right:
    st.subheader("运行日志")
    progress_bar = st.progress(0)
    log_container = st.empty()
    log_container.markdown(
        '<div class="log-box"><div class="log-entry">等待任务开始。</div></div>',
        unsafe_allow_html=True,
    )

    if st.button("开始逐段润色", use_container_width=True, type="primary"):
        if not uploaded_file:
            st.error("请先上传 Word 文档。")
        elif not api_key.startswith("sk-"):
            st.error("请填写有效的 DeepSeek API Key。")
        else:
            st.session_state.processed_file = None
            progress_bar.progress(0)
            result_bytes = process_word(
                uploaded_file.read(),
                api_key,
                model_name,
                concurrency,
                start_page,
                end_page,
                min_chars,
                prompt,
                log_container,
                progress_bar,
            )

            if result_bytes:
                st.session_state.processed_file = result_bytes

    if st.session_state.processed_file:
        st.success("处理完成，可以下载新文档。")
        st.download_button(
            label="下载润色后的 Word 文档",
            data=st.session_state.processed_file,
            file_name=f"润色版_{uploaded_file.name if uploaded_file else 'document.docx'}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )

if st.session_state.rewrite_report:
    st.divider()
    st.subheader("改写对照清单")

    successful_count = sum(1 for item in st.session_state.rewrite_report if item["status"] == "success")
    failed_count = len(st.session_state.rewrite_report) - successful_count
    metric_col_1, metric_col_2, metric_col_3 = st.columns(3)
    metric_col_1.metric("处理段落", len(st.session_state.rewrite_report))
    metric_col_2.metric("成功替换", successful_count)
    metric_col_3.metric("保留原文", failed_count)

    download_col_1, download_col_2 = st.columns(2)
    download_col_1.download_button(
        "下载对照清单 JSON",
        data=report_to_json(st.session_state.rewrite_report),
        file_name="rewrite_report.json",
        mime="application/json",
        use_container_width=True,
    )
    download_col_2.download_button(
        "下载对照清单 CSV",
        data=report_to_csv(st.session_state.rewrite_report),
        file_name="rewrite_report.csv",
        mime="text/csv",
        use_container_width=True,
    )

    sorted_report = sorted(
        st.session_state.rewrite_report,
        key=lambda item: (item["page"], item["paragraph_index"]),
    )
    for item in sorted_report:
        status_text = "已替换" if item["status"] == "success" else "保留原文"
        with st.expander(f"第 {item['page']} 页 / 第 {item['paragraph_index']} 段 - {status_text}"):
            before_col, after_col = st.columns(2)
            before_col.markdown("**原文**")
            before_col.write(item["original_text"])
            after_col.markdown("**改写后**")
            after_col.write(item["new_text"])

            if item["status"] == "success":
                st.markdown("**差异高亮**")
                st.markdown(
                    f'<div class="diff-box">{item["diff_html"]}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.warning(f"这一段没有写回，原因：{item['error']}")
