import concurrent.futures
import csv
import difflib
import html
import io
import json
import re
import shutil
import time
import uuid
import zipfile
from pathlib import Path
from html.parser import HTMLParser
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
HEADING_LABEL_RE = re.compile(r"^(heading|标题|標題)\s*([12])$", re.IGNORECASE)
SKIP_STYLE_KEYWORDS = (
    "title",
    "heading",
    "toc",
    "caption",
    "header",
    "footer",
    "bibliography",
    "reference",
)
TITLE_LIKE_RE = re.compile(
    r"^(abstract|chapter|section|table of contents|contents|references|bibliography|appendix|acknowledgements?|摘要|关键词|關鍵詞|目录|目錄)\b",
    re.IGNORECASE,
)
KEYWORD_LINE_RE = re.compile(r"^\s*(keywords?|关键词|關鍵詞)\s*[:：]", re.IGNORECASE)
TITLE_CASE_WORD_RE = re.compile(r"^[A-Z][A-Za-z'’.-]*(?:\s+[A-Z][A-Za-z'’.-]*)*$")
ACADEMIC_TITLE_HINT_RE = re.compile(
    r"\b(A\s+Corpus-Based\s+Study|Case\s+Study|Translation\s+of|Research\s+on|Study\s+on)\b|"
    r"(基于.+研究|以.+为例|翻译研究)",
    re.IGNORECASE,
)
PROTECTED_FRONT_MATTER_RE = re.compile(
    r"^(毕业论文原创性声明|论文作者签名|日期：|本人郑重声明|Acknowledgements?|I would like to express|"
    r"In Partial Fulfillment|Under the Supervision|School of Foreign Studies|Nanjing University)",
    re.IGNORECASE,
)
CASUAL_REWRITE_RE = re.compile(
    r"\b(kicked off|packed with|zeroes in|let'?s cut|does not explain themselves|nipped at the heels|"
    r"upended the field|hits its limits|doesn'?t|don'?t|isn'?t|aren'?t|won'?t|can'?t|the result\?)\b",
    re.IGNORECASE,
)
PROMPT_TEMPLATES = {
    "中国学生 GRE 代入法": """你是一个中国的学生，你刚刚通过了gre考试，作文的分数为3，这是你的水平。请你针对以下的英文内容进行重写，想尽力写出大概4分左右水平的英文作文，请记住你是一个中国大学生，可能词汇量和语言习惯和其他国家的学生不同，请你切实的代入后进行写作，减少ai风格的语言习惯，切实代入人类学生，减少模版写作的痕迹，在写作时加入一定的个人色彩。可以在语言上显示出瑕疵，但是你的学术态度端正，请你展现出这种不完美和瑕疵感。注意保留原有的专业术语、数字和引用，不要随意增加或删除客观事实。仅返回重写后的段落文本，不要添加任何Markdown格式（如星号等）重要:原本是中文的千万不要变成英文""",
    "降 AI 率（英文）": """You are a professional AIGC-rate reduction tool. Rewrite the paragraph according to the following rules. The primary goal is to reduce AIGC-like writing while preserving academic professionalism.

Core requirements:
- You may adjust sentence structure and use a slightly more natural, mildly conversational rhythm, but do not lose the professional thesis tone.
- All proper nouns, technical terms, names, titles, citations, numbers, data, and terminology must be preserved completely and accurately.
- If a sentence contains many proper nouns or technical terms and is difficult to restructure, simplify it and compress its length.
- If a sentence contains few proper nouns or technical terms, expand it moderately and make the expression more detailed.
- Keep the central meaning unchanged. Do not add unsupported facts, data, citations, or conclusions.
- Use synonymous expressions, change sentence openings, and vary clause order where possible.
- For titles, keyword lines, references, captions, and table/list labels, return the original unchanged.
- Do not use Markdown formatting such as *italics* or **bold**.

Return only the rewritten paragraph.""",
    "降 AI 率（中文）": """请重写下面文本，使表达更自然、更少模板感和机器生成痕迹。避免使用“综上所述”“此外”“进一步而言”等套话式衔接，尽量改用更贴合上下文的过渡方式。调整句式开头和句长，多用具体动词，减少抽象名词堆叠。必要时拆分长句，并进行实质性改写，不要原样返回。""",
    "学术润色（英文）": """Rewrite the text with clearer, plainer scholarly prose while preserving the original claim, evidence, citations, terminology, numbers, and order of information. Remove formulaic transitions and inflated academic phrasing, but do not make the paragraph conversational, punchy, metaphorical, or editorialized. Do not add new content. Return only the rewritten paragraph.""",
    "深度自然化（英文）": """Rewrite the paragraph as a conservative thesis revision. Preserve the original meaning, facts, citations, numbers, and technical terms. Change sentence openings, clause order, and some wording where needed, but keep the register plain and academic.

Avoid both AI-polished prose and AI-naturalized prose:
- no contractions, rhetorical questions, idioms, jokes, or punchy fragments;
- no magazine-style verbs such as "kicked off," "packed with," "zeroes in," or "digs into";
- no first-person unless the original paragraph already requires it;
- no added examples or explanatory claims.
- no Markdown formatting such as *italics* or **bold**.

Return only the rewritten paragraph.""",
    "深度降 AI（英文强改写）": """Rewrite this paragraph more decisively while preserving the same claim, evidence, citations, terminology, numbers, and order of information. The target is restrained undergraduate thesis prose, not polished editor prose and not casual humanization.

Important style target:
- Do not produce glossy, generic, perfectly balanced academic prose.
- Do not produce conversational or magazine-like prose.
- Avoid formulaic connectors, broad concluding language, and repeated sentence frames.
- Remove filler and stock phrasing instead of replacing it with new stock phrasing.
- Use concrete verbs and specific links to the local context.
- Vary sentence length and syntax modestly; some sentences may be plain and compact.
- If the source is wordy, compress it. If the source is thin, do not add new detail.
- Do not use contractions, rhetorical questions, idioms, jokes, punchy fragments, or first-person unless present in the source.
- Do not use Markdown formatting such as *italics* or **bold**.

Make substantial structural and lexical changes, not just synonym swaps. Return only the rewritten paragraph.""",
}


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
        border: 1px solid #374151;
        border-radius: 8px;
        padding: 12px;
        background: #111827;
        color: #e5e7eb;
        line-height: 1.7;
        min-height: 44px;
    }
    .diff-del {
        color: #991b1b;
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
    .diff-replace-old {
        color: #7f1d1d;
        background: #fecaca;
        padding: 1px 3px;
        border-radius: 4px;
    }
    .diff-replace-new {
        color: #92400e;
        background: #fef3c7;
        padding: 1px 3px;
        border-radius: 4px;
    }
    .compare-row {
        border: 1px solid #374151;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 14px;
        background: #111827;
    }
    .compare-meta {
        font-weight: 700;
        margin-bottom: 12px;
    }
    .compare-text {
        line-height: 1.7;
        white-space: pre-wrap;
        word-break: break-word;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "logs" not in st.session_state:
    st.session_state.logs = []
if "processed_file" not in st.session_state:
    st.session_state.processed_file = None
if "tracked_file" not in st.session_state:
    st.session_state.tracked_file = None
if "rewrite_report" not in st.session_state:
    st.session_state.rewrite_report = []
if "output_prefix" not in st.session_state:
    st.session_state.output_prefix = "润色版"


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


def element_xml(element):
    if element is None:
        return ""
    return ET.tostring(element, encoding="unicode")


def clone_element(element):
    if element is None:
        return None
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def is_pure_text_run(run):
    if run.tag != f"{W_NS}r":
        return False
    text_children = 0
    for child in list(run):
        if child.tag == f"{W_NS}rPr":
            continue
        if child.tag != f"{W_NS}t":
            return False
        text_children += 1
    return text_children > 0


def run_style_key(run):
    return element_xml(run.find(f"{W_NS}rPr"))


def run_text_nodes(run):
    return [child for child in list(run) if child.tag == f"{W_NS}t"]


def merge_adjacent_text_runs(parent):
    merged_count = 0
    children = list(parent)
    idx = 0
    while idx < len(children) - 1:
        current = children[idx]
        nxt = children[idx + 1]
        if (
            is_pure_text_run(current)
            and is_pure_text_run(nxt)
            and run_style_key(current) == run_style_key(nxt)
        ):
            current_text_nodes = run_text_nodes(current)
            next_text_nodes = run_text_nodes(nxt)
            if current_text_nodes and next_text_nodes:
                target = current_text_nodes[-1]
                appended_text = "".join(node.text or "" for node in next_text_nodes)
                target.text = (target.text or "") + appended_text
                if target.text.startswith(" ") or target.text.endswith(" "):
                    target.set(XML_SPACE, "preserve")
                parent.remove(nxt)
                children.pop(idx + 1)
                merged_count += 1
                continue
        idx += 1

    for child in list(parent):
        merged_count += merge_adjacent_text_runs(child)
    return merged_count


def make_diff_html_pair(old_text, new_text):
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    old_parts = []
    new_parts = []

    for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        old_part = " ".join(old_words[old_start:old_end])
        new_part = " ".join(new_words[new_start:new_end])

        if tag == "equal":
            old_parts.append(html.escape(old_part))
            new_parts.append(html.escape(new_part))
        elif tag == "delete":
            old_parts.append(f'<del class="diff-del">{html.escape(old_part)}</del>')
        elif tag == "insert":
            new_parts.append(f'<ins class="diff-ins">{html.escape(new_part)}</ins>')
        elif tag == "replace":
            old_parts.append(
                f'<span class="diff-replace-old" title="改写后：{html.escape(new_part, quote=True)}">'
                f"{html.escape(old_part)}</span>"
            )
            new_parts.append(
                f'<span class="diff-replace-new" title="原文：{html.escape(old_part, quote=True)}">'
                f"{html.escape(new_part)}</span>"
            )

    return {
        "old_html": " ".join(part for part in old_parts if part),
        "new_html": " ".join(part for part in new_parts if part),
    }


def report_to_json(report):
    return json.dumps(report, ensure_ascii=False, indent=2).encode("utf-8")


def report_to_csv(report):
    output = io.StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=[
            "paragraph_index",
            "page",
            "section",
            "section_type",
            "write_mode",
            "rewrite_rounds",
            "second_rewrite_applied",
            "reject_reason",
            "repair_mode",
            "match_score",
            "matched_fragments",
            "risk_tags",
            "final_risk_tags",
            "original_text",
            "new_text",
            "status",
            "error",
        ],
    )
    writer.writeheader()
    for item in report:
        writer.writerow(
            {
                "paragraph_index": item.get("paragraph_index", ""),
                "page": item.get("page", ""),
                "section": item.get("section", ""),
                "section_type": item.get("section_type", ""),
                "write_mode": item.get("write_mode", ""),
                "rewrite_rounds": item.get("rewrite_rounds", ""),
                "second_rewrite_applied": item.get("second_rewrite_applied", ""),
                "reject_reason": item.get("reject_reason", ""),
                "repair_mode": item.get("repair_mode", ""),
                "match_score": item.get("match_score", ""),
                "matched_fragments": item.get("matched_fragments", ""),
                "risk_tags": item.get("risk_tags", ""),
                "final_risk_tags": item.get("final_risk_tags", ""),
                "original_text": item.get("original_text", ""),
                "new_text": item.get("new_text", ""),
                "status": item.get("status", ""),
                "error": item.get("error", ""),
            }
        )
    return output.getvalue().encode("utf-8-sig")


def split_visible_paragraphs(text):
    return [line.strip() for line in text.splitlines()]


def paragraph_ref_only(paragraph_listing):
    return paragraph_listing.split("|", 1)[0].strip()


def generate_tracked_changes_docx(original_bytes, revised_bytes):
    from docx_editor import Document

    work_dir = Path("tmp") / "redline" / uuid.uuid4().hex
    work_dir.mkdir(parents=True, exist_ok=False)
    original_path = work_dir / "original.docx"
    revised_path = work_dir / "revised.docx"
    output_path = work_dir / "tracked.docx"

    original_path.write_bytes(original_bytes)
    revised_path.write_bytes(revised_bytes)

    revised_doc = None
    original_doc = None
    try:
        revised_doc = Document.open(revised_path, author="AI Refiner", force_recreate=True)
        revised_paragraphs = split_visible_paragraphs(revised_doc.get_visible_text())
        revised_doc.close(cleanup=False)
        revised_doc = None

        original_doc = Document.open(original_path, author="AI Refiner", force_recreate=True)
        original_paragraphs = split_visible_paragraphs(original_doc.get_visible_text())
        paragraph_refs = [
            paragraph_ref_only(item) for item in original_doc.list_paragraphs(max_chars=0)
        ]

        count = min(len(original_paragraphs), len(revised_paragraphs), len(paragraph_refs))
        rewrites = []
        for idx in range(count):
            old_text = original_paragraphs[idx]
            new_text = revised_paragraphs[idx]
            if old_text and old_text != new_text:
                rewrites.append((paragraph_refs[idx], new_text))

        if rewrites:
            original_doc.batch_rewrite(rewrites)
        original_doc.save(output_path)
        original_doc.close(cleanup=False)
        original_doc = None

        return output_path.read_bytes(), len(rewrites)
    finally:
        if revised_doc is not None:
            revised_doc.close(cleanup=False)
        if original_doc is not None:
            original_doc.close(cleanup=False)
        try:
            shutil.rmtree(work_dir)
        except Exception:
            pass


def make_system_prompt(user_prompt, extra_instruction=""):
    retry_instruction = f"\n\n{extra_instruction}" if extra_instruction else ""
    return f"""{user_prompt}{retry_instruction}

请严格遵守：
1. 只返回改写后的正文段落。
2. 不要添加解释、Markdown、代码块、标题或前后缀。
3. 不要添加原文没有依据的新事实、数据、文献、引文或结论。
4. 必须保持与输入文本相同的主体语言。若输入主要是英文，改写后也必须是英文；不要把段落翻译成另一种语言，除非用户明确要求翻译。中文术语、标题、人名或引文可以按原样保留。
5. 不要把文本改成过度工整、过度平滑、处处对称的 AI 润色腔。宁可保留适度自然的不均匀节奏，也不要生成模板化套话。
6. 如果原文有具体上下文，优先用上下文内的衔接方式，不要使用空泛连接词或泛泛总结句。
"""


class SuspiciousSpanParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.depth = 0
        self.current = []
        self.fragments = []

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        class_value = attrs_dict.get("class", "")
        if tag in ("a", "span", "div") and re.search(r"\b(cl[123]|hide_3)\b", class_value):
            self.depth += 1
        elif self.depth:
            self.depth += 1

    def handle_endtag(self, tag):
        if not self.depth:
            return
        self.depth -= 1
        if self.depth == 0:
            text = normalize_visible_text("".join(self.current))
            if text:
                self.fragments.append(text)
            self.current = []

    def handle_data(self, data):
        if self.depth:
            self.current.append(data)


def normalize_visible_text(text):
    return re.sub(r"\s+", " ", html.unescape(text or "")).strip()


def normalize_for_match(text):
    text = html.unescape(text or "").lower()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^\w\u4e00-\u9fff]+", "", text)
    return text


def clean_html_fragment(raw_html):
    text = re.sub(r"<br\s*/?>", " ", raw_html, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    return normalize_visible_text(text)


def extract_paperyy_aigc_fragments_from_html(text):
    fragments = []
    pattern = re.compile(
        r"<em\b(?=[^>]*class=['\"][^'\"]*\b(?:high|medium|low)\b[^'\"]*['\"])[^>]*>(.*?)</em>",
        re.I | re.S,
    )
    for match in pattern.finditer(text):
        fragments.append(clean_html_fragment(match.group(1)))
    return fragments


def dedupe_fragments(fragments, min_chars=28):
    seen = set()
    results = []
    for fragment in fragments:
        clean = normalize_visible_text(fragment)
        if len(clean) < min_chars:
            continue
        key = normalize_for_match(clean)
        if len(key) < min_chars or key in seen:
            continue
        seen.add(key)
        results.append(clean)
    return results


def extract_report_fragments_from_html(report_bytes):
    text = report_bytes.decode("utf-8", errors="ignore")
    fragments = []

    parser = SuspiciousSpanParser()
    try:
        parser.feed(text)
        fragments.extend(parser.fragments)
    except Exception:
        pass

    for match in re.finditer(r"<span[^>]*class=['\"]hide_3['\"][^>]*>(.*?)</span>", text, re.I | re.S):
        fragments.append(clean_html_fragment(match.group(1)))
    for match in re.finditer(r"<a[^>]*class=['\"][^'\"]*\bcl[123]\b[^'\"]*['\"][^>]*>(.*?)</a>", text, re.I | re.S):
        fragments.append(clean_html_fragment(match.group(1)))
    fragments.extend(extract_paperyy_aigc_fragments_from_html(text))

    return dedupe_fragments(fragments)


def extract_report_fragments_from_pdf(report_bytes):
    try:
        from pypdf import PdfReader
    except Exception as exc:
        raise RuntimeError("PDF 报告解析需要 pypdf 依赖，请先安装 requirements.txt。") from exc

    reader = PdfReader(io.BytesIO(report_bytes))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    candidates = []
    for line in text.splitlines():
        clean = normalize_visible_text(line)
        if len(clean) >= 40:
            candidates.append(clean)
    return dedupe_fragments(candidates)


def extract_report_fragments(report_bytes, report_name):
    lower_name = (report_name or "").lower()
    if lower_name.endswith(".html") or lower_name.endswith(".htm"):
        return extract_report_fragments_from_html(report_bytes)
    if lower_name.endswith(".pdf"):
        return extract_report_fragments_from_pdf(report_bytes)
    raise RuntimeError("检测报告只支持 HTML、HTM 或 PDF。")


def match_report_fragments_to_tasks(fragments, tasks, max_tasks=None):
    matches_by_task = {}
    normalized_tasks = [(task, normalize_for_match(task["plain_text"])) for task in tasks]

    for fragment in fragments:
        norm_fragment = normalize_for_match(fragment)
        if len(norm_fragment) < 28:
            continue
        best_task = None
        best_score = 0.0
        for task, norm_text in normalized_tasks:
            if not norm_text:
                continue
            score = 0.0
            if norm_fragment in norm_text:
                score = min(1.0, len(norm_fragment) / max(len(norm_text), 1) + 0.35)
            elif norm_text in norm_fragment:
                score = min(0.95, len(norm_text) / max(len(norm_fragment), 1))
            elif len(norm_fragment) >= 80:
                score = difflib.SequenceMatcher(None, norm_fragment, norm_text).ratio()
            if score > best_score:
                best_score = score
                best_task = task

        if best_task is None or best_score < 0.58:
            continue
        key = best_task["paragraph_index"]
        entry = matches_by_task.setdefault(
            key,
            {
                "task": best_task,
                "fragments": [],
                "score": 0.0,
            },
        )
        entry["score"] = max(entry["score"], best_score)
        if len(entry["fragments"]) < 5:
            entry["fragments"].append(fragment)

    matches = sorted(
        matches_by_task.values(),
        key=lambda item: (item["task"]["page"], item["task"]["paragraph_index"]),
    )
    if max_tasks:
        matches = matches[:max_tasks]
    return matches


def call_deepseek(
    text,
    user_prompt,
    api_key,
    model_name,
    task_id,
    temperature=0.4,
    extra_instruction="",
    max_retries=3,
):
    url = "https://api.deepseek.com/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": make_system_prompt(user_prompt, extra_instruction)},
            {"role": "user", "content": text},
        ],
        "temperature": temperature,
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


def call_deepseek_direct(
    text,
    system_prompt,
    api_key,
    model_name,
    task_id,
    temperature=0.4,
    max_retries=3,
):
    url = "https://api.deepseek.com/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text},
        ],
        "temperature": temperature,
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

def call_baidu_translate(text, from_lang, to_lang, appid, secret_key):
    import hashlib
    salt = str(int(__import__('time').time() * 1000))
    sign_str = appid + text + salt + secret_key
    sign = hashlib.md5(sign_str.encode("utf-8")).hexdigest()
    params = {"q": text, "from": from_lang, "to": to_lang,
              "appid": appid, "salt": salt, "sign": sign}
    resp = requests.get(
        "https://fanyi-api.baidu.com/api/trans/vip/translate",
        params=params, timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if "error_code" in data:
        raise RuntimeError(f"百度翻译错误 {data['error_code']}: {data.get('error_msg', '')}")
    return " ".join(item["dst"] for item in data["trans_result"])


def call_baidu_chain(text, lang_chain, baidu_appid, baidu_secret_key):
    current_text = text
    for step_i in range(len(lang_chain) - 1):
        from_lang = lang_chain[step_i]
        to_lang = lang_chain[step_i + 1]
        current_text = call_baidu_translate(
            current_text, from_lang, to_lang, baidu_appid, baidu_secret_key
        )
        time.sleep(1.1)
    return current_text



def direct_text_runs(paragraph):
    return [child for child in list(paragraph) if child.tag == f"{W_NS}r" and is_pure_text_run(child)]


def paragraph_direct_text_nodes(paragraph):
    nodes = []
    for run in direct_text_runs(paragraph):
        for text_elem in run_text_nodes(run):
            if text_elem.text:
                nodes.append(text_elem)
    return nodes


def paragraph_all_text_nodes(paragraph):
    return [elem for elem in paragraph.iter(f"{W_NS}t") if elem.text]


def has_complex_inline_content(paragraph):
    for child in list(paragraph):
        if child.tag == f"{W_NS}pPr":
            continue
        if child.tag == f"{W_NS}r" and is_pure_text_run(child):
            continue
        return True
    return False


def run_format_flags(run):
    r_pr = run.find(f"{W_NS}rPr")
    if r_pr is None:
        return {}
    flags = {}
    if r_pr.find(f"{W_NS}i") is not None:
        flags["italic"] = True
    if r_pr.find(f"{W_NS}b") is not None:
        flags["bold"] = True
    if r_pr.find(f"{W_NS}u") is not None:
        flags["underline"] = True
    return flags


def formatted_terms_from_paragraph(paragraph):
    terms = {}
    for run in direct_text_runs(paragraph):
        flags = run_format_flags(run)
        if not flags:
            continue
        text = "".join(node.text or "" for node in run_text_nodes(run)).strip()
        if len(text) < 2 or not re.search(r"[\w\u4e00-\u9fff]", text):
            continue
        existing = terms.setdefault(text, {})
        existing.update(flags)
    return [{"text": text, "flags": flags} for text, flags in terms.items()]


def paragraph_plain_text(paragraph):
    return "".join(elem.text or "" for elem in paragraph.iter(f"{W_NS}t")).strip()


def word_count(text):
    return len(re.findall(r"[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*|[\u4e00-\u9fff]", text))


def ends_with_sentence_punctuation(text):
    return text.rstrip().endswith((".", "?", "!", "。", "？", "！", ";", "；", ":"))


def paragraph_style_id(paragraph):
    p_pr = paragraph.find(f"{W_NS}pPr")
    if p_pr is None:
        return ""
    p_style = p_pr.find(f"{W_NS}pStyle")
    if p_style is None:
        return ""
    return p_style.get(f"{W_NS}val", "")


def read_style_names(doc_zip):
    style_names = {}
    try:
        styles_root = ET.fromstring(doc_zip.read("word/styles.xml"))
    except Exception:
        return style_names

    for style in styles_root.findall(f"{W_NS}style"):
        style_id = style.get(f"{W_NS}styleId", "")
        name = style.find(f"{W_NS}name")
        if style_id and name is not None:
            style_names[style_id] = name.get(f"{W_NS}val", "")
    return style_names


def paragraph_style_label(paragraph, style_names):
    style_id = paragraph_style_id(paragraph)
    style_name = style_names.get(style_id, "")
    return f"{style_id} {style_name}".strip()


def heading_level(paragraph, style_names):
    label = paragraph_style_label(paragraph, style_names).replace("_", " ").strip()
    compact_label = re.sub(r"\s+", "", label).lower()
    if "heading1" in compact_label or "标题1" in compact_label or "標題1" in compact_label:
        return 1
    if "heading2" in compact_label or "标题2" in compact_label or "標題2" in compact_label:
        return 2

    match = HEADING_LABEL_RE.match(re.sub(r"\s+", " ", label).strip())
    if match:
        return int(match.group(2))
    return None


def extract_headings_from_docx(file_bytes):
    headings = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as doc_zip:
            root = ET.fromstring(doc_zip.read("word/document.xml"))
            style_names = read_style_names(doc_zip)
    except Exception:
        return headings

    for idx, paragraph in enumerate(root.iter(f"{W_NS}p")):
        level = heading_level(paragraph, style_names)
        if level not in (1, 2):
            continue
        text = paragraph_plain_text(paragraph)
        if not text:
            continue
        heading_no = len(headings) + 1
        headings.append(
            {
                "label": f"{heading_no}. H{level} - {text[:80]}",
                "paragraph_index": idx,
                "level": level,
                "text": text,
            }
        )
    return headings


def heading_section_end(headings, heading_position):
    if heading_position is None:
        return None
    current = headings[heading_position]
    for next_heading in headings[heading_position + 1 :]:
        if next_heading["level"] <= current["level"]:
            return next_heading["paragraph_index"]
    return None


def paragraph_page_number(paragraph, current_page):
    page_breaks = len(paragraph.findall(f".//{W_NS}lastRenderedPageBreak"))
    for br in paragraph.findall(f".//{W_NS}br"):
        if br.get(f"{W_NS}type") == "page":
            page_breaks += 1
    return current_page + page_breaks


def is_body_paragraph(paragraph, style_names=None):
    style_names = style_names or {}
    text = paragraph_plain_text(paragraph)
    if not text:
        return False
    style_label = paragraph_style_label(paragraph, style_names).lower()
    if any(keyword in style_label for keyword in SKIP_STYLE_KEYWORDS):
        return False
    compact_text = re.sub(r"\s+", " ", text).strip()
    if should_skip_protected_text(compact_text):
        return False
    words = word_count(compact_text)
    if words < 12:
        return False
    if words < 20 and not ends_with_sentence_punctuation(compact_text):
        return False
    return True


def looks_like_academic_title(text):
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return False
    if ends_with_sentence_punctuation(compact):
        return False
    words = compact.split()
    if len(words) > 32:
        return False
    if ACADEMIC_TITLE_HINT_RE.search(compact):
        return True
    if re.search(r"[\u4e00-\u9fff]", compact) and not re.search(r"[。！？.!?]", compact):
        return bool(ACADEMIC_TITLE_HINT_RE.search(compact)) or (12 <= len(compact) <= 90)
    titleish_words = sum(1 for word in words if word[:1].isupper())
    return len(words) >= 8 and titleish_words / max(len(words), 1) >= 0.55


def should_skip_protected_text(text):
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return True
    if KEYWORD_LINE_RE.match(compact):
        return True
    if PROTECTED_FRONT_MATTER_RE.match(compact):
        return True
    if TITLE_LIKE_RE.match(compact) and len(compact) <= 120:
        return True
    if looks_like_academic_title(compact):
        return True
    return False


def collect_tasks(root, style_names, start_paragraph_index, end_paragraph_index, min_chars):
    tasks = []
    current_page = 1
    current_heading = ""
    current_section_type = "general"
    end_index = end_paragraph_index if end_paragraph_index is not None else float("inf")

    paragraphs = list(root.iter(f"{W_NS}p"))
    for idx, paragraph in enumerate(paragraphs):
        current_page = paragraph_page_number(paragraph, current_page)
        level = heading_level(paragraph, style_names)
        if level in (1, 2):
            current_heading = paragraph_plain_text(paragraph)
            current_section_type = detect_section_type(current_heading)
            continue
        paragraph_text = paragraph_plain_text(paragraph)
        compact_paragraph_text = re.sub(r"\s+", " ", paragraph_text).strip()
        if re.match(r"^(abstract|摘\s*要)\b", compact_paragraph_text, re.IGNORECASE):
            current_heading = compact_paragraph_text
            current_section_type = "abstract"
        elif re.match(r"^chapter\s+one\b", compact_paragraph_text, re.IGNORECASE):
            current_heading = compact_paragraph_text
            current_section_type = "introduction"
        elif re.match(r"^(references?|bibliography|参考文献)\b", compact_paragraph_text, re.IGNORECASE):
            current_heading = compact_paragraph_text
            current_section_type = "references"
        if idx <= start_paragraph_index or idx >= end_index:
            continue
        if current_section_type == "references":
            continue
        if not is_body_paragraph(paragraph, style_names):
            continue
        text_runs = direct_text_runs(paragraph)
        has_complex_content = has_complex_inline_content(paragraph)
        text_nodes = (
            paragraph_all_text_nodes(paragraph)
            if has_complex_content
            else paragraph_direct_text_nodes(paragraph)
        )
        plain_text = "".join(elem.text or "" for elem in text_nodes).strip()
        if len(plain_text) < min_chars:
            continue

        tasks.append(
            {
                "paragraph_index": idx,
                "page": current_page,
                "p_node": paragraph,
                "text_runs": text_runs,
                "text_nodes": text_nodes,
                "plain_text": plain_text,
                "write_mode": "text_nodes_fallback" if has_complex_content else "rebuild_runs",
                "section_heading": current_heading,
                "section_type": current_section_type,
                "risk_profile": analyze_ai_risk(plain_text, current_section_type),
                "formatted_terms": formatted_terms_from_paragraph(paragraph),
            }
        )
    return tasks


def ensure_run_format(r_pr, flags):
    if flags.get("bold") and r_pr.find(f"{W_NS}b") is None:
        ET.SubElement(r_pr, f"{W_NS}b")
    if flags.get("italic") and r_pr.find(f"{W_NS}i") is None:
        ET.SubElement(r_pr, f"{W_NS}i")
    if flags.get("underline") and r_pr.find(f"{W_NS}u") is None:
        underline = ET.SubElement(r_pr, f"{W_NS}u")
        underline.set(f"{W_NS}val", "single")


def make_text_run(text, base_r_pr=None, flags=None):
    run = ET.Element(f"{W_NS}r")
    r_pr = clone_element(base_r_pr)
    flags = flags or {}
    if flags:
        if r_pr is None:
            r_pr = ET.Element(f"{W_NS}rPr")
        ensure_run_format(r_pr, flags)
    if r_pr is not None:
        run.append(r_pr)
    text_elem = ET.SubElement(run, f"{W_NS}t")
    if text.startswith(" ") or text.endswith(" "):
        text_elem.set(XML_SPACE, "preserve")
    text_elem.text = text
    return run


def split_text_by_terms(text, formatted_terms):
    terms = sorted(
        [term for term in formatted_terms if term["text"] and term["text"] in text],
        key=lambda term: len(term["text"]),
        reverse=True,
    )
    if not terms:
        return [(text, {})]

    pattern = re.compile("|".join(re.escape(term["text"]) for term in terms))
    flags_by_text = {term["text"]: term["flags"] for term in terms}
    pieces = []
    last = 0
    for match in pattern.finditer(text):
        if match.start() > last:
            pieces.append((text[last : match.start()], {}))
        matched_text = match.group(0)
        pieces.append((matched_text, flags_by_text.get(matched_text, {})))
        last = match.end()
    if last < len(text):
        pieces.append((text[last:], {}))
    return [(piece, flags) for piece, flags in pieces if piece]


def rewrite_paragraph_text(task, new_text):
    if task.get("write_mode") == "text_nodes_fallback":
        text_nodes = task.get("text_nodes", [])
        if not text_nodes:
            raise ValueError("Paragraph has no writable text nodes.")
        first_node = text_nodes[0]
        first_node.text = new_text
        if new_text.startswith(" ") or new_text.endswith(" "):
            first_node.set(XML_SPACE, "preserve")
        elif XML_SPACE in first_node.attrib:
            first_node.attrib.pop(XML_SPACE, None)
        for node in text_nodes[1:]:
            node.text = ""
            node.attrib.pop(XML_SPACE, None)
        return

    paragraph = task["p_node"]
    runs = task["text_runs"]
    if not runs:
        raise ValueError("段落没有可写入的纯文本 run。")

    children = list(paragraph)
    first_index = children.index(runs[0])
    base_r_pr = runs[0].find(f"{W_NS}rPr")
    for run in runs:
        paragraph.remove(run)

    new_runs = [
        make_text_run(piece, base_r_pr, flags)
        for piece, flags in split_text_by_terms(new_text, task["formatted_terms"])
    ]
    for offset, run in enumerate(new_runs):
        paragraph.insert(first_index + offset, run)


def suspicious_rewrite_reason(original_text, new_text, enforce_format_safety=True):
    original_words = word_count(original_text)
    new_words = word_count(new_text)
    original_compact = re.sub(r"\s+", " ", original_text).strip()
    new_compact = re.sub(r"\s+", " ", new_text).strip()
    if should_skip_protected_text(original_compact) and original_compact != new_compact:
        return "protected_text_changed"
    if enforce_format_safety and "*" not in original_compact and "*" in new_compact:
        return "markdown_formatting_added"
    if enforce_format_safety and CASUAL_REWRITE_RE.search(new_compact):
        return "casual_rewrite_phrase"
    if not ends_with_sentence_punctuation(original_compact) and ends_with_sentence_punctuation(new_compact):
        if new_words > max(20, int(original_words * (2.0 if not enforce_format_safety else 1.6))):
            return "title_or_label_expanded"
    if original_words < 20 and new_words > max(30, int(original_words * (3.5 if not enforce_format_safety else 2.5))):
        return "short_text_expanded"
    if original_words >= 30 and new_words > int(original_words * (2.5 if not enforce_format_safety else 1.45)):
        return "rewrite_too_long"
    return ""


def is_suspicious_expansion(original_text, new_text):
    return bool(suspicious_rewrite_reason(original_text, new_text))


AI_RISK_RULES = (
    {
        "tag": "generic_study_opening",
        "pattern": r"\b(This study|The findings|The results|Existing research|Previous studies|It is important to note|To build upon|This research|This thesis)\b",
        "instruction": "Replace generic study/report openings with a direct claim tied to this paragraph's local content.",
    },
    {
        "tag": "template_connectors",
        "pattern": r"\b(moreover|furthermore|in addition|in conclusion|overall|by contrast|on the other hand|from the perspective of|according to)\b",
        "instruction": "Remove template connectors unless they are necessary; use a specific transition based on the surrounding argument.",
    },
    {
        "tag": "stock_interpretive_verbs",
        "pattern": r"\b(shows|reveals|demonstrates|highlights|indicates|suggests|reflects|underscores|illustrates|exhibits|presents|provides|offers)\b",
        "instruction": "Avoid repeatedly using stock interpretive verbs such as shows/reveals/suggests/presents; state the relation more concretely.",
    },
    {
        "tag": "future_research_template",
        "pattern": r"\b(future studies should|further research|future scholars|subsequent research|future work|longitudinal analysis would|future studies|future research)\b",
        "instruction": "Rewrite future-research language as specific unanswered questions or concrete next steps, not a formulaic closing paragraph.",
    },
    {
        "tag": "results_data_template",
        "pattern": r"\b(Tables?\s+\d|Figures?\s+\d|First,\s+regarding|percentage-point gap|frequency distribution|quantitative distribution|version exhibits|version shows|was adopted in|identified CLWs)\b",
        "instruction": "Rewrite table/result reporting so it does not follow the fixed 'table presents + percentages + broad implication' pattern.",
    },
    {
        "tag": "contribution_template",
        "pattern": r"\b(contributions? that extend|replicable analytical pathway|methodological reference|theoretical and practical implications|offers several contributions|contribute to a deeper understanding)\b",
        "instruction": "Rewrite contribution claims as restrained, specific takeaways; avoid broad contribution formulas.",
    },
    {
        "tag": "over_casual_naturalization",
        "pattern": r"\b(real headaches|make or break|isn't a coincidence|sticks closer|leans into|zeroes in|on the ground|kicked off|packed with|digs into|let'?s cut|upended the field|hits its limits|nipped at the heels)\b",
        "instruction": "Avoid casual AI-naturalized phrasing, idioms, and punchy editorial language; keep the wording plain and thesis-appropriate.",
    },
    {
        "tag": "first_person_method",
        "pattern": r"\b(I|we|my|our)\b",
        "instruction": "Avoid first-person narration unless the original paragraph already used it; keep methodology procedural and thesis-appropriate.",
    },
)


ABSTRACT_NOUN_RE = re.compile(
    r"\b(tendency|strategy|framework|approach|distribution|analysis|difference|divergence|context|purpose|orientation|significance|methodology|classification|category|pattern|translation|foreignisation|foreignization|domestication)\b",
    re.IGNORECASE,
)
DATA_RE = re.compile(r"\d+(?:\.\d+)?\s*%|\b\d+(?:\.\d+)?\b")


def detect_section_type(heading_text):
    heading = (heading_text or "").lower()
    if re.search(r"\b(references?|bibliography)\b|参考文献", heading, re.IGNORECASE):
        return "references"
    if "abstract" in heading:
        return "abstract"
    if "introduction" in heading:
        return "introduction"
    if "literature" in heading or "review" in heading:
        return "literature_review"
    if "method" in heading or "methodology" in heading:
        return "methodology"
    if "result" in heading or "discussion" in heading or "analysis" in heading:
        return "results_discussion"
    if "conclusion" in heading:
        return "conclusion"
    return "general"


def analyze_ai_risk(text, section_type="general"):
    tags = []
    instructions = []
    for rule in AI_RISK_RULES:
        if re.search(rule["pattern"], text, re.IGNORECASE):
            tags.append(rule["tag"])
            instructions.append(rule["instruction"])

    abstract_nouns = len(ABSTRACT_NOUN_RE.findall(text))
    if abstract_nouns >= 6:
        tags.append("abstract_noun_density")
        instructions.append(
            "Reduce abstract noun stacking; turn at least some strategy/framework/approach language into concrete actions, evidence, or relations."
        )

    has_data = bool(DATA_RE.search(text))
    if has_data and re.search(
        r"\b(show|shows|reveal|reveals|suggest|suggests|indicate|indicates|reflect|reflects|account|accounts|represent|represents|present|presents|exhibit|exhibits|adopt|adopts|adopted)\b",
        text,
        re.IGNORECASE,
    ):
        tags.append("formulaic_data_explanation")
        instructions.append(
            "For numbers or percentages, avoid the fixed pattern 'data + shows/suggests + broad meaning'; explain the specific comparison in plainer terms."
        )

    section_instructions = {
        "abstract": "This appears to be abstract-like writing: keep it compact, factual, and less formulaic; avoid a chain of 'This study...' sentences.",
        "introduction": "This appears to be introduction-like writing: reduce broad background claims and move quickly to the concrete research gap.",
        "literature_review": "This appears to be literature-review writing: keep citations intact, but avoid the repeated 'existing studies have... however...' template.",
        "methodology": "This appears to be methodology writing: preserve reproducible steps, but do not turn it into a polished narrative with unnecessary first person.",
        "results_discussion": "This appears to be results/discussion writing: keep figures and comparisons accurate, but vary how the paragraph interprets data.",
        "conclusion": "This appears to be conclusion-like writing: avoid generic contribution and future-study formulas; make the limitations and next steps specific.",
    }
    if section_type in section_instructions:
        instructions.insert(0, section_instructions[section_type])
        tags.append(f"section_{section_type}")

    return {
        "tags": sorted(set(tags)),
        "score": len(set(tags)) + min(abstract_nouns // 6, 2),
        "instructions": instructions,
    }


def build_risk_instruction(risk_profile):
    instructions = risk_profile.get("instructions", [])
    if not instructions:
        return ""
    bullets = "\n".join(f"- {item}" for item in instructions[:8])
    return f"""Paragraph-specific risk repair:
{bullets}

Apply these repairs while preserving meaning, factual claims, citations, terminology, and the paragraph's role in the thesis."""


HIGH_PRIORITY_RISK_TAGS = {
    "formulaic_data_explanation",
    "results_data_template",
    "future_research_template",
    "contribution_template",
    "generic_study_opening",
    "abstract_noun_density",
}


def task_temperature(base_temperature, risk_profile, section_type):
    tags = set(risk_profile.get("tags", []))
    boost = 0
    if tags & HIGH_PRIORITY_RISK_TAGS:
        boost += 0.06
    if section_type in ("abstract", "results_discussion", "conclusion"):
        boost += 0.05
    if risk_profile.get("score", 0) >= 4:
        boost += 0.04
    return min(base_temperature + boost, 0.88)


def should_force_risk_rewrite(original_text, new_text, before_risk, after_risk, section_type):
    before_tags = set(before_risk.get("tags", []))
    after_tags = set(after_risk.get("tags", []))
    non_section_after_tags = {tag for tag in after_tags if not tag.startswith("section_")}
    high_after = after_tags & HIGH_PRIORITY_RISK_TAGS
    similarity = difflib.SequenceMatcher(None, original_text, new_text).ratio()

    if similarity >= 0.92 and before_tags & HIGH_PRIORITY_RISK_TAGS:
        return True
    if high_after:
        return True
    if section_type in ("results_discussion", "conclusion") and non_section_after_tags:
        return True
    if before_risk.get("score", 0) >= 4 and after_risk.get("score", 0) >= 2:
        return True
    return False


def build_force_rewrite_instruction(before_risk, after_risk, section_type):
    combined = {
        "instructions": list(before_risk.get("instructions", []))
        + list(after_risk.get("instructions", []))
    }
    section_focus = {
        "abstract": "For an abstract, remove formulaic purpose/result phrasing and keep the paragraph compact, direct, and information-dense.",
        "results_discussion": "For a results/discussion paragraph, keep every figure accurate but rewrite the explanation around the specific comparison instead of using 'shows/reveals/suggests'.",
        "conclusion": "For a conclusion paragraph, remove broad contribution and future-research formulas; state limits, implications, and next steps in specific terms.",
    }.get(section_type, "Rewrite the paragraph more decisively while keeping the thesis-appropriate register.")

    return (
        build_risk_instruction(combined)
        + "\n\nMandatory second-pass rewrite:\n"
        + f"- {section_focus}\n"
        + "- Do not preserve sentence order if it keeps the same formulaic rhythm.\n"
        + "- Replace any remaining generic study openings, table-reporting formulas, and future-study stock phrases.\n"
        + "- Keep citations, numbers, names, and technical terms exact."
    )



def process_word(
    file_bytes,
    api_key,
    model_name,
    concurrency,
    start_paragraph_index,
    end_paragraph_index,
    min_chars,
    prompt,
    rewrite_temperature,
    adaptive_risk_repair,
    enforce_format_safety,
    generate_tracked_file,
    log_container,
    progress_bar,
    progress_status=None,
):
    st.session_state.logs = []
    st.session_state.rewrite_report = []
    st.session_state.tracked_file = None

    def update_progress(value, message):
        safe_value = min(max(float(value), 0.0), 1.0)
        progress_bar.progress(safe_value)
        if progress_status is not None:
            progress_status.caption(message)

    try:
        update_progress(0.01, "正在读取并解析 Word 文档...")
        add_log("正在读取并解析 Word 文档。")
        render_logs(log_container)

        input_buffer = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_buffer, "r") as doc_zip:
            xml_content = doc_zip.read("word/document.xml")
            root = ET.fromstring(xml_content)
            style_names = read_style_names(doc_zip)
            merged_count = merge_adjacent_text_runs(root)

            add_log(f"已合并 {merged_count} 个相邻同样式纯文本 run，正在扫描正文段落。")
            tasks = collect_tasks(root, style_names, start_paragraph_index, end_paragraph_index, min_chars)
            if not tasks:
                add_log("在设定范围内没有找到符合长度要求的正文段落。", "warn")
                render_logs(log_container)
                update_progress(0, "没有找到可处理段落。")
                return None

            add_log(f"共找到 {len(tasks)} 个待处理段落，开始 {concurrency} 并发处理。", "info")
            update_progress(0.06, f"已找到 {len(tasks)} 个待处理段落，正在提交给模型...")
            fallback_count = sum(
                1 for task in tasks if task.get("write_mode") == "text_nodes_fallback"
            )
            if fallback_count:
                add_log(
                    f"{fallback_count} paragraphs contain complex inline structure and will use fallback text-node writing.",
                    "warn",
                )
            render_logs(log_container)

            completed_tasks = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=concurrency) as executor:
                future_to_task = {}
                for task in tasks:
                    risk_instruction = (
                        build_risk_instruction(task["risk_profile"]) if adaptive_risk_repair else ""
                    )
                    task["risk_instruction"] = risk_instruction
                    current_temperature = (
                        task_temperature(
                            rewrite_temperature,
                            task["risk_profile"],
                            task["section_type"],
                        )
                        if adaptive_risk_repair
                        else rewrite_temperature
                    )
                    task["rewrite_temperature"] = current_temperature
                    future = executor.submit(
                        call_deepseek,
                        task["plain_text"],
                        prompt,
                        api_key,
                        model_name,
                        task["paragraph_index"] + 1,
                        current_temperature,
                        risk_instruction,
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
                    rewrite_rounds = ""
                    second_rewrite_applied = False
                    reject_reason = ""
                    try:
                        response_text = future.result()
                        original_text = task["plain_text"]
                        new_text = response_text.strip()
                        retry_error = ""
                        rewrite_rounds = 1
                        second_rewrite_applied = False
                        reject_reason = ""
                        if original_text == new_text:
                            try:
                                retry_response_text = call_deepseek(
                                    task["plain_text"],
                                    prompt,
                                    api_key,
                                    model_name,
                                    para_no,
                                    temperature=min(task.get("rewrite_temperature", rewrite_temperature) + 0.1, 0.9),
                                    extra_instruction=(
                                        task.get("risk_instruction", "")
                                        + "\n\n"
                                        "Your previous revision was identical to the original. "
                                        "Revise again with more substantial wording and sentence-level changes, "
                                        "while preserving meaning, scholarly tone, and factual boundaries."
                                    ),
                                    max_retries=2,
                                )
                                retry_new_text = retry_response_text.strip()
                                if retry_new_text != original_text:
                                    new_text = retry_new_text
                                    rewrite_rounds += 1
                                    add_log(f"第 {para_no} 段首次无变化，已自动重试并改写。", "info")
                            except Exception as retry_exc:
                                retry_error = str(retry_exc)
                                add_log(
                                    f"第 {para_no} 段无变化重试失败，保留第一次合法结果：{retry_exc}",
                                    "warn",
                                )
                        elif adaptive_risk_repair:
                            before_risk = task["risk_profile"]
                            after_risk = analyze_ai_risk(new_text, task["section_type"])
                            if should_force_risk_rewrite(
                                original_text,
                                new_text,
                                before_risk,
                                after_risk,
                                task["section_type"],
                            ):
                                try:
                                    retry_response_text = call_deepseek(
                                        task["plain_text"],
                                        prompt,
                                        api_key,
                                        model_name,
                                        para_no,
                                        temperature=min(
                                            task.get("rewrite_temperature", rewrite_temperature) + 0.14,
                                            0.92,
                                        ),
                                        extra_instruction=(
                                            build_force_rewrite_instruction(
                                                before_risk,
                                                after_risk,
                                                task["section_type"],
                                            )
                                            + "\n\nYour previous rewrite still retained too many detector-prone patterns. "
                                            "Rewrite again with less formulaic structure, fewer stock transitions, "
                                            "and more paragraph-specific phrasing."
                                        ),
                                        max_retries=2,
                                    )
                                    retry_new_text = retry_response_text.strip()
                                    if retry_new_text and retry_new_text != original_text:
                                        new_text = retry_new_text
                                        rewrite_rounds += 1
                                        second_rewrite_applied = True
                                        add_log(
                                            f"第 {para_no} 段仍有检测器敏感模式，已执行第二轮定向改写。",
                                            "info",
                                        )
                                except Exception as retry_exc:
                                    retry_error = str(retry_exc)
                                    add_log(
                                        f"Paragraph {para_no}: second risk-repair rewrite failed; keeping the first result: {retry_exc}",
                                        "warn",
                                    )
                        diff_html = make_diff_html_pair(original_text, new_text)
                        status = "changed" if original_text != new_text else "unchanged"
                        final_risk = analyze_ai_risk(new_text, task["section_type"])
                        reject_reason = suspicious_rewrite_reason(
                            original_text,
                            new_text,
                            enforce_format_safety=enforce_format_safety,
                        )
                        if reject_reason:
                            raise ValueError(f"改写结果触发安全阀：{reject_reason}，已拒绝写回。")
                        rewrite_paragraph_text(task, new_text)
                        st.session_state.rewrite_report.append(
                            {
                                "paragraph_index": para_no,
                                "page": task["page"],
                                "original_text": original_text,
                                "new_text": new_text,
                                "old_diff_html": diff_html["old_html"],
                                "new_diff_html": diff_html["new_html"],
                                "status": status,
                                "error": retry_error if status == "unchanged" else "",
                                "section": task["section_heading"],
                                "section_type": task["section_type"],
                                "write_mode": task.get("write_mode", ""),
                                "rewrite_rounds": rewrite_rounds,
                                "second_rewrite_applied": second_rewrite_applied,
                                "reject_reason": reject_reason,
                                "risk_tags": ", ".join(task["risk_profile"]["tags"]),
                                "final_risk_tags": ", ".join(final_risk["tags"]),
                            }
                        )
                        if status == "changed":
                            add_log(f"第 {para_no} 段处理完成，并已写回原文本节点。", "success")
                        else:
                            add_log(f"第 {para_no} 段处理完成，但文本无明显变化。", "info")
                    except Exception as exc:
                        st.session_state.rewrite_report.append(
                            {
                                "paragraph_index": para_no,
                                "page": task["page"],
                                "original_text": task["plain_text"],
                                "new_text": task["plain_text"],
                                "old_diff_html": "",
                                "new_diff_html": "",
                                "status": "failed",
                                "error": str(exc),
                                "section": task.get("section_heading", ""),
                                "section_type": task.get("section_type", ""),
                                "write_mode": task.get("write_mode", ""),
                                "rewrite_rounds": rewrite_rounds,
                                "second_rewrite_applied": second_rewrite_applied,
                                "reject_reason": reject_reason,
                                "risk_tags": ", ".join(task.get("risk_profile", {}).get("tags", [])),
                                "final_risk_tags": "",
                            }
                        )
                        add_log(f"第 {para_no} 段处理失败，已保留原文：{exc}", "err")

                    completed_tasks += 1
                    rewrite_progress = 0.06 + (completed_tasks / len(tasks)) * 0.84
                    update_progress(
                        rewrite_progress,
                        f"AI 改写中：已完成 {completed_tasks}/{len(tasks)} 段。高风险段落可能会自动二次改写。",
                    )
                    render_logs(log_container)

            update_progress(0.92, "改写完成，正在打包润色版 Word 文档...")
            output_io = io.BytesIO()
            with zipfile.ZipFile(output_io, "w", zipfile.ZIP_DEFLATED) as out_zip:
                for item in doc_zip.infolist():
                    if item.filename != "word/document.xml":
                        out_zip.writestr(item, doc_zip.read(item.filename))

                xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                out_zip.writestr("word/document.xml", xml_str)

        clean_docx = output_io.getvalue()
        add_log("润色版文档已打包完成。原段落、run 和样式结构没有被重建。", "success")
        if generate_tracked_file:
            update_progress(0.96, "润色版已生成，正在生成修订痕迹版...")
            try:
                tracked_docx, tracked_count = generate_tracked_changes_docx(file_bytes, clean_docx)
                st.session_state.tracked_file = tracked_docx
                add_log(f"修订版文档已生成，共写入 {tracked_count} 个段落级修订。", "success")
            except Exception as tracked_exc:
                add_log(f"修订版生成失败，已保留润色版下载：{tracked_exc}", "warn")
        else:
            add_log("已跳过修订痕迹版生成，处理速度更快。", "info")
        render_logs(log_container)
        update_progress(1.0, "处理完成，可以下载新文档。")
        return clean_docx

    except Exception as exc:
        add_log(f"处理流程发生异常：{exc}", "err")
        render_logs(log_container)
        update_progress(0, "处理失败，请查看运行日志。")
        return None


def process_word_baidu(
    file_bytes,
    baidu_appid,
    baidu_secret_key,
    lang_chain,
    concurrency,
    start_paragraph_index,
    end_paragraph_index,
    min_chars,
    enforce_format_safety,
    generate_tracked_file,
    log_container,
    progress_bar,
    progress_status=None,
):
    st.session_state.logs = []
    st.session_state.rewrite_report = []
    st.session_state.tracked_file = None

    def update_progress(value, message):
        safe_value = min(max(float(value), 0.0), 1.0)
        progress_bar.progress(safe_value)
        if progress_status is not None:
            progress_status.caption(message)

    try:
        update_progress(0.01, "正在读取并解析 Word 文档...")
        add_log("正在读取并解析 Word 文档。")
        render_logs(log_container)

        input_buffer = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_buffer, "r") as doc_zip:
            xml_content = doc_zip.read("word/document.xml")
            root = ET.fromstring(xml_content)
            style_names = read_style_names(doc_zip)
            merged_count = merge_adjacent_text_runs(root)

            add_log(f"已合并 {merged_count} 个相邻同样式纯文本 run，正在扫描正文段落。")
            tasks = collect_tasks(root, style_names, start_paragraph_index, end_paragraph_index, min_chars)
            if not tasks:
                add_log("在设定范围内没有找到符合长度要求的正文段落。", "warn")
                render_logs(log_container)
                update_progress(0, "没有找到可处理段落。")
                return None

            chain_label = " → ".join(lang_chain)
            steps = len(lang_chain) - 1
            add_log(f"共找到 {len(tasks)} 个待处理段落，翻译链条：{chain_label}（共 {steps} 步）。", "info")
            add_log(f"注意：百度翻译免费版 QPS=1，每段约需 {steps * 1.1:.0f}~{steps * 2:.0f} 秒，请耐心等待。", "warn")
            update_progress(0.06, f"已找到 {len(tasks)} 个待处理段落，开始回译...")
            render_logs(log_container)

            completed_tasks = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=concurrency) as executor:
                future_to_task = {}
                for task in tasks:
                    future = executor.submit(
                        call_baidu_chain,
                        task["plain_text"],
                        lang_chain,
                        baidu_appid,
                        baidu_secret_key
                    )
                    future_to_task[future] = task
                    add_log(f"第 {task['paragraph_index'] + 1} 段，第 {task['page']} 页，已加入多语言翻译队列。", "send")
                    render_logs(log_container)

                for future in concurrent.futures.as_completed(future_to_task):
                    task = future_to_task[future]
                    para_no = task["paragraph_index"] + 1
                    original_text = task["plain_text"]
                    reject_reason = ""
                    try:
                        response_text = future.result()
                        new_text = response_text.strip()
                        diff_html = make_diff_html_pair(original_text, new_text)
                        status = "changed" if original_text != new_text else "unchanged"
                        reject_reason = suspicious_rewrite_reason(
                            original_text, new_text, enforce_format_safety=enforce_format_safety
                        )
                        if reject_reason:
                            raise ValueError(f"改写结果触发安全阀：{reject_reason}，已拒绝写回。")
                        rewrite_paragraph_text(task, new_text)
                        final_risk = analyze_ai_risk(new_text, task["section_type"])
                        st.session_state.rewrite_report.append({
                            "paragraph_index": para_no,
                            "page": task["page"],
                            "original_text": original_text,
                            "new_text": new_text,
                            "old_diff_html": diff_html["old_html"],
                            "new_diff_html": diff_html["new_html"],
                            "status": status,
                            "error": "",
                            "section": task["section_heading"],
                            "section_type": task["section_type"],
                            "write_mode": task.get("write_mode", ""),
                            "rewrite_rounds": len(lang_chain) - 1,
                            "second_rewrite_applied": False,
                            "reject_reason": "",
                            "risk_tags": ", ".join(task["risk_profile"]["tags"]),
                            "final_risk_tags": ", ".join(final_risk["tags"]),
                        })
                        if status == "changed":
                            add_log(f"第 {para_no} 段回译完成，已写回。", "success")
                        else:
                            add_log(f"第 {para_no} 段回译完成，但文本无明显变化。", "info")
                    except Exception as exc:
                        st.session_state.rewrite_report.append({
                            "paragraph_index": para_no,
                            "page": task["page"],
                            "original_text": original_text,
                            "new_text": original_text,
                            "old_diff_html": "",
                            "new_diff_html": "",
                            "status": "failed",
                            "error": str(exc),
                            "section": task.get("section_heading", ""),
                            "section_type": task.get("section_type", ""),
                            "write_mode": task.get("write_mode", ""),
                            "rewrite_rounds": "",
                            "second_rewrite_applied": False,
                            "reject_reason": reject_reason,
                            "risk_tags": ", ".join(task.get("risk_profile", {}).get("tags", [])),
                            "final_risk_tags": "",
                        })
                        add_log(f"第 {para_no} 段回译失败，已保留原文：{exc}", "err")

                    completed_tasks += 1
                    update_progress(
                        0.06 + completed_tasks / len(tasks) * 0.84,
                        f"回译中：已完成 {completed_tasks}/{len(tasks)} 段。"
                    )
                    render_logs(log_container)

            update_progress(0.92, "回译完成，正在打包 Word 文档...")
            output_io = io.BytesIO()
            with zipfile.ZipFile(output_io, "w", zipfile.ZIP_DEFLATED) as out_zip:
                for item in doc_zip.infolist():
                    if item.filename != "word/document.xml":
                        out_zip.writestr(item, doc_zip.read(item.filename))
                xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                out_zip.writestr("word/document.xml", xml_str)

        clean_docx = output_io.getvalue()
        add_log("回译版文档已打包完成。", "success")
        if generate_tracked_file:
            update_progress(0.96, "回译版已生成，正在生成修订痕迹版...")
            try:
                tracked_docx, tracked_count = generate_tracked_changes_docx(file_bytes, clean_docx)
                st.session_state.tracked_file = tracked_docx
                add_log(f"修订版文档已生成，共写入 {tracked_count} 个段落级修订。", "success")
            except Exception as tracked_exc:
                add_log(f"修订版生成失败，已保留回译版下载：{tracked_exc}", "warn")
        else:
            add_log("已跳过修订痕迹版生成，处理速度更快。", "info")
        render_logs(log_container)
        update_progress(1.0, "回译完成，可以下载新文档。")
        return clean_docx

    except Exception as exc:
        add_log(f"回译流程发生异常：{exc}", "err")
        render_logs(log_container)
        update_progress(0, "回译失败，请查看运行日志。")
        return None


def process_report_repair_word(
    file_bytes,
    report_bytes,
    report_name,
    tracked_base_bytes,
    api_key,
    model_name,
    concurrency,
    min_chars,
    prompt,
    rewrite_temperature,
    adaptive_risk_repair,
    enforce_format_safety,
    generate_tracked_file,
    log_container,
    progress_bar,
    progress_status=None,
):
    st.session_state.logs = []
    st.session_state.rewrite_report = []
    st.session_state.tracked_file = None

    def update_progress(value, message):
        safe_value = min(max(float(value), 0.0), 1.0)
        progress_bar.progress(safe_value)
        if progress_status is not None:
            progress_status.caption(message)

    try:
        update_progress(0.01, "正在读取检测报告...")
        fragments = extract_report_fragments(report_bytes, report_name)
        add_log(f"检测报告中提取到 {len(fragments)} 个疑似片段。", "info")
        render_logs(log_container)
        if not fragments:
            add_log("未能从检测报告中提取疑似 AI 片段，请优先上传 HTML 统计报告。", "warn")
            update_progress(0, "未提取到疑似片段。")
            return None

        update_progress(0.05, "正在解析 Word 并匹配疑似片段...")
        input_buffer = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_buffer, "r") as doc_zip:
            xml_content = doc_zip.read("word/document.xml")
            root = ET.fromstring(xml_content)
            style_names = read_style_names(doc_zip)
            merged_count = merge_adjacent_text_runs(root)
            tasks = collect_tasks(root, style_names, -1, None, min_chars)
            matches = match_report_fragments_to_tasks(fragments, tasks)
            if not matches:
                add_log("疑似片段没有匹配到可改写的 Word 正文段落。", "warn")
                update_progress(0, "未匹配到可返修段落。")
                render_logs(log_container)
                return None

            add_log(
                f"已合并 {merged_count} 个相邻 run。检测报告命中 {len(matches)} 个 Word 段落，开始定点返修。",
                "info",
            )
            render_logs(log_container)
            update_progress(0.10, f"已匹配 {len(matches)} 个高风险段落，正在提交给模型...")

            completed_tasks = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=concurrency) as executor:
                future_to_match = {}
                for match in matches:
                    task = match["task"]
                    matched_text = "\n".join(f"- {frag[:500]}" for frag in match["fragments"][:3])
                    report_instruction = (
                        "This paragraph was matched from an AIGC detector report and needs targeted repair.\n"
                        "Detector-flagged fragment(s):\n"
                        f"{matched_text}\n\n"
                        "Revise the whole paragraph, but focus on removing the detector-prone rhythm in those fragments. "
                        "Keep citations, names, numbers, terminology, and claims exact. Do not add new content."
                    )
                    risk_instruction = (
                        build_risk_instruction(task["risk_profile"]) if adaptive_risk_repair else ""
                    )
                    task["risk_instruction"] = "\n\n".join(
                        item for item in (report_instruction, risk_instruction) if item
                    )
                    current_temperature = (
                        task_temperature(
                            rewrite_temperature,
                            task["risk_profile"],
                            task["section_type"],
                        )
                        if adaptive_risk_repair
                        else rewrite_temperature
                    )
                    task["rewrite_temperature"] = min(current_temperature + 0.04, 0.9)
                    future = executor.submit(
                        call_deepseek,
                        task["plain_text"],
                        prompt,
                        api_key,
                        model_name,
                        task["paragraph_index"] + 1,
                        task["rewrite_temperature"],
                        task["risk_instruction"],
                    )
                    future_to_match[future] = match
                    add_log(
                        f"第 {task['paragraph_index'] + 1} 段命中检测报告，已发送给模型。",
                        "send",
                    )

                render_logs(log_container)
                for future in concurrent.futures.as_completed(future_to_match):
                    match = future_to_match[future]
                    task = match["task"]
                    para_no = task["paragraph_index"] + 1
                    rewrite_rounds = ""
                    second_rewrite_applied = False
                    reject_reason = ""
                    try:
                        response_text = future.result()
                        original_text = task["plain_text"]
                        new_text = response_text.strip()
                        retry_error = ""
                        rewrite_rounds = 1

                        before_risk = task["risk_profile"]
                        after_risk = analyze_ai_risk(new_text, task["section_type"])
                        if adaptive_risk_repair and should_force_risk_rewrite(
                            original_text,
                            new_text,
                            before_risk,
                            after_risk,
                            task["section_type"],
                        ):
                            try:
                                retry_response_text = call_deepseek(
                                    task["plain_text"],
                                    prompt,
                                    api_key,
                                    model_name,
                                    para_no,
                                    temperature=min(task.get("rewrite_temperature", rewrite_temperature) + 0.12, 0.92),
                                    extra_instruction=(
                                        task.get("risk_instruction", "")
                                        + "\n\nMandatory report-repair second pass: the first rewrite still looks detector-prone. "
                                        "Make a more substantial but restrained thesis-style revision. Keep all factual content exact."
                                    ),
                                    max_retries=2,
                                )
                                retry_new_text = retry_response_text.strip()
                                if retry_new_text and retry_new_text != original_text:
                                    new_text = retry_new_text
                                    rewrite_rounds += 1
                                    second_rewrite_applied = True
                                    add_log(f"第 {para_no} 段返修后仍有风险，已执行第二轮定向改写。", "info")
                            except Exception as retry_exc:
                                retry_error = str(retry_exc)
                                add_log(f"第 {para_no} 段返修二次改写失败，保留第一轮结果：{retry_exc}", "warn")

                        diff_html = make_diff_html_pair(original_text, new_text)
                        status = "changed" if original_text != new_text else "unchanged"
                        final_risk = analyze_ai_risk(new_text, task["section_type"])
                        reject_reason = suspicious_rewrite_reason(
                            original_text,
                            new_text,
                            enforce_format_safety=enforce_format_safety,
                        )
                        if reject_reason:
                            raise ValueError(f"改写结果触发安全阀：{reject_reason}，已拒绝写回。")
                        rewrite_paragraph_text(task, new_text)
                        st.session_state.rewrite_report.append(
                            {
                                "paragraph_index": para_no,
                                "page": task["page"],
                                "original_text": original_text,
                                "new_text": new_text,
                                "old_diff_html": diff_html["old_html"],
                                "new_diff_html": diff_html["new_html"],
                                "status": status,
                                "error": retry_error if status == "unchanged" else "",
                                "section": task["section_heading"],
                                "section_type": task["section_type"],
                                "write_mode": task.get("write_mode", ""),
                                "rewrite_rounds": rewrite_rounds,
                                "second_rewrite_applied": second_rewrite_applied,
                                "reject_reason": reject_reason,
                                "repair_mode": "aigc_report",
                                "match_score": f"{match['score']:.3f}",
                                "matched_fragments": " || ".join(match["fragments"][:5]),
                                "risk_tags": ", ".join(task["risk_profile"]["tags"]),
                                "final_risk_tags": ", ".join(final_risk["tags"]),
                            }
                        )
                        add_log(f"第 {para_no} 段定点返修完成。", "success")
                    except Exception as exc:
                        st.session_state.rewrite_report.append(
                            {
                                "paragraph_index": para_no,
                                "page": task["page"],
                                "original_text": task["plain_text"],
                                "new_text": task["plain_text"],
                                "old_diff_html": "",
                                "new_diff_html": "",
                                "status": "failed",
                                "error": str(exc),
                                "section": task.get("section_heading", ""),
                                "section_type": task.get("section_type", ""),
                                "write_mode": task.get("write_mode", ""),
                                "rewrite_rounds": rewrite_rounds,
                                "second_rewrite_applied": second_rewrite_applied,
                                "reject_reason": reject_reason,
                                "repair_mode": "aigc_report",
                                "match_score": f"{match['score']:.3f}",
                                "matched_fragments": " || ".join(match["fragments"][:5]),
                                "risk_tags": ", ".join(task.get("risk_profile", {}).get("tags", [])),
                                "final_risk_tags": "",
                            }
                        )
                        add_log(f"第 {para_no} 段定点返修失败，已保留原文：{exc}", "err")

                    completed_tasks += 1
                    update_progress(
                        0.10 + (completed_tasks / len(matches)) * 0.80,
                        f"检测报告返修中：已完成 {completed_tasks}/{len(matches)} 段。",
                    )
                    render_logs(log_container)

            update_progress(0.92, "定点返修完成，正在打包 Word 文档...")
            output_io = io.BytesIO()
            with zipfile.ZipFile(output_io, "w", zipfile.ZIP_DEFLATED) as out_zip:
                for item in doc_zip.infolist():
                    if item.filename != "word/document.xml":
                        out_zip.writestr(item, doc_zip.read(item.filename))
                xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                out_zip.writestr("word/document.xml", xml_str)

        clean_docx = output_io.getvalue()
        add_log("检测报告返修版文档已打包完成。", "success")
        if generate_tracked_file:
            tracked_base = tracked_base_bytes or file_bytes
            tracked_label = "原始第一版" if tracked_base_bytes else "当前上传文档"
            update_progress(0.96, f"返修版已生成，正在生成相对{tracked_label}的修订痕迹版...")
            try:
                tracked_docx, tracked_count = generate_tracked_changes_docx(tracked_base, clean_docx)
                st.session_state.tracked_file = tracked_docx
                add_log(
                    f"修订版文档已生成，比较基准：{tracked_label}，共写入 {tracked_count} 个段落级修订。",
                    "success",
                )
            except Exception as tracked_exc:
                add_log(f"修订版生成失败，已保留返修版下载：{tracked_exc}", "warn")
        else:
            add_log("已跳过修订痕迹版生成，处理速度更快。", "info")
        render_logs(log_container)
        update_progress(1.0, "检测报告返修完成，可以下载新文档。")
        return clean_docx
    except Exception as exc:
        add_log(f"检测报告返修流程发生异常：{exc}", "err")
        render_logs(log_container)
        update_progress(0, "检测报告返修失败，请查看运行日志。")
        return None
st.title("AI Word 论文逐段润色工具")
st.markdown("上传 .docx 后，系统会逐段发送正文给 AI，并只替换 Word XML 里的文本节点，尽量保持原有格式结构不变。")

col_left, col_right = st.columns([1, 1.2])

with col_left:
    st.subheader("1. 处理模式")
    process_mode = st.radio(
        "处理模式",
        ["整篇逐段润色", "检测报告定点返修", "中文中转回译（百度翻译）"],
        horizontal=True,
        help="定点返修只改 AIGC 检测报告命中的段落；回译模式使用百度翻译 API，无需 DeepSeek。",
    )

    st.subheader("2. API 配置")
    if process_mode != "中文中转回译（百度翻译）":
        api_key = st.text_input("DeepSeek API Key", type="password")
        col_1, col_2 = st.columns(2)
        model_name = col_1.selectbox("模型", ["deepseek-chat", "deepseek-reasoner"])
        concurrency = col_2.slider(
            "并发请求数", min_value=1, max_value=8, value=3,
            help="建议 2-4。并发过高可能触发 API 限流。",
        )
        baidu_appid = ""
        baidu_secret_key = ""
    else:
        api_key = ""
        model_name = "deepseek-chat"
        col_b1, col_b2 = st.columns(2)
        baidu_appid = col_b1.text_input("百度翻译 APPID", key="baidu_appid")
        baidu_secret_key = col_b2.text_input("百度翻译密钥（Secret Key）", type="password", key="baidu_secret_key")
        concurrency = st.slider(
            "并发请求数", min_value=1, max_value=10, value=1,
            help="百度翻译免费版 QPS 限制为 1，必须设置为 1。如果使用高级版（QPS=10），可调高并发加速处理。",
        )
        st.caption(
            "在 [百度翻译开放平台](https://fanyi-api.baidu.com/) 免费注册，"
            "免费版 QPS=1，必须单线程；高级版 QPS=10，可开 10 线程。"
        )

    st.subheader("3. 上传文档")
    uploaded_file = st.file_uploader("选择 Word 文档 (.docx)", type=["docx"])
    generate_tracked_file = st.checkbox(
        "生成修订痕迹版（较慢）",
        value=False,
        help="关闭后只生成润色/返修后的 Word，速度更快；需要对比修改痕迹时再开启。",
    )
    aigc_report_file = None
    original_base_file = None
    if process_mode == "检测报告定点返修":
        aigc_report_file = st.file_uploader(
            "上传 AIGC 检测报告（HTML/PDF）",
            type=["html", "htm", "pdf"],
            help="优先上传检测平台导出的 HTML 统计报告；PDF 也可尝试解析。",
        )
        if generate_tracked_file:
            original_base_file = st.file_uploader(
                "可选：上传第一版/原始 Word，用于生成相对原文的修订版",
                type=["docx"],
                help='不上传则修订版默认比较当前待返修 Word -> 返修版。',
            )

    headings = extract_headings_from_docx(uploaded_file.getvalue()) if uploaded_file else []
    heading_labels = [heading["label"] for heading in headings]
    heading_lookup = {heading["label"]: idx for idx, heading in enumerate(headings)}

    st.subheader("4. 章节范围")
    if headings:
        start_heading_label = st.selectbox("从哪个章节开始润色", ["全文开头"] + heading_labels)
        end_heading_label = st.selectbox("到哪个章节结束", ["全文末尾"] + heading_labels)
        start_heading_pos = heading_lookup.get(start_heading_label)
        end_heading_pos = heading_lookup.get(end_heading_label)
        start_paragraph_index = headings[start_heading_pos]["paragraph_index"] if start_heading_pos is not None else -1
        end_paragraph_index = heading_section_end(headings, end_heading_pos) if end_heading_pos is not None else None
        if end_heading_pos is not None and headings[end_heading_pos]["paragraph_index"] <= start_paragraph_index:
            st.warning("结束章节位于开始章节之前，当前范围可能没有可处理正文。")
    else:
        start_paragraph_index = -1
        end_paragraph_index = None
        st.info("未识别到 Heading 1/2，将默认处理全文正文。")

    min_chars = st.slider("忽略短段落，少于 N 字符不处理", min_value=5, max_value=120, value=30, step=5)

    if process_mode != "中文中转回译（百度翻译）":
        st.subheader("5. 润色提示词")
        prompt_template_name = st.selectbox("提示词模板", list(PROMPT_TEMPLATES.keys()))
        prompt = st.text_area("提示词", value=PROMPT_TEMPLATES[prompt_template_name], height=180)
        rewrite_strength = st.selectbox(
            "改写强度",
            ["标准降 AI", "深度降 AI", "最大强改写"],
            index=1,
            help="强度越高，句式和措辞变化越大；如果检测结果仍偏高，优先尝试【最大强改写】。",
        )
        temperature_by_strength = {"标准降 AI": 0.55, "深度降 AI": 0.68, "最大强改写": 0.78}
        rewrite_temperature = temperature_by_strength[rewrite_strength]
        adaptive_risk_repair = st.checkbox(
            "启用段落风险扫描与自动二次修复",
            value=True,
            help="自动识别模板句式、抽象名词堆叠和数据解释套话，并把这些问题传给模型定向改写。",
        )
    else:
        prompt = ""
        rewrite_temperature = 0.55
        adaptive_risk_repair = False
        st.subheader("5. 多语言回译链条")
        BAIDU_LANG_OPTIONS = {
            "英语 (en)": "en",
            "中文 (zh)": "zh",
            "法语 (fra)": "fra",
            "俄语 (ru)": "ru",
            "日语 (jp)": "jp",
            "德语 (de)": "de",
            "西班牙语 (spa)": "spa",
            "阿拉伯语 (ara)": "ara",
            "韩语 (kor)": "kor",
            "葡萄牙语 (pt)": "pt",
        }
        PRESET_CHAINS = {
            "英→中→英（2步，最快）": ["en", "zh", "en"],
            "英→法→中→英（3步）": ["en", "fra", "zh", "en"],
            "英→法→中→俄→英（4步）": ["en", "fra", "zh", "ru", "en"],
            "英→日→法→中→英（4步）": ["en", "jp", "fra", "zh", "en"],
            "英→德→日→中→俄→英（5步）": ["en", "de", "jp", "zh", "ru", "en"],
            "自定义": None,
        }
        preset_name = st.selectbox(
            "预设链条",
            list(PRESET_CHAINS.keys()),
            index=1,
            help="选择预设翻译链条，步数越多混淆效果越强，但耗时更长。",
        )
        if PRESET_CHAINS[preset_name] is not None:
            lang_chain = PRESET_CHAINS[preset_name]
            chain_display = " → ".join(
                next(k for k, v in BAIDU_LANG_OPTIONS.items() if v == c)
                for c in lang_chain
            )
            st.caption(f"当前链条：{chain_display}")
        else:
            st.caption("自定义中间语言（起点英语 en，终点英语 en 固定，选择中间经过的语言）")
            mid_langs = st.multiselect(
                "中间语言（按顺序选择）",
                [k for k in BAIDU_LANG_OPTIONS if k != "英语 (en)"],
                default=["法语 (fra)", "中文 (zh)"],
                key="custom_mid_langs",
            )
            if not mid_langs:
                st.warning("请至少选择一种中间语言。")
                mid_langs = ["中文 (zh)"]
            lang_chain = ["en"] + [BAIDU_LANG_OPTIONS[k] for k in mid_langs] + ["en"]
            chain_display = " → ".join(
                next(k for k, v in BAIDU_LANG_OPTIONS.items() if v == c)
                for c in lang_chain
            )
            st.caption(f"当前链条：{chain_display}")
        steps = len(lang_chain) - 1
        st.info(
            f"翻译步数：{steps} 步，每段约需 {steps * 1.1:.0f}~{steps * 2:.0f} 秒（免费版 QPS=1）。\n"
            "所有格式安全阀、长度保护、protected text 检测均正常生效。"
        )

    enforce_format_safety = st.checkbox(
        "启用格式安全阀",
        value=True,
        help="开启时，禁止输出新增 Markdown 星号或过于随意的口语化词组，并严格限制段落扩写长度（最大 1.45 倍）；"
             "关闭后不仅允许星号和口语化，还会将字数扩写限制大幅放宽（最高允许 2.5 倍），以适应大量废话扩写。",
    )

with col_right:
    st.subheader("运行日志")
    progress_bar = st.progress(0)
    progress_status = st.empty()
    progress_status.caption("等待任务开始。")
    log_container = st.empty()
    log_container.markdown(
        '<div class="log-box"><div class="log-entry">等待任务开始。</div></div>',
        unsafe_allow_html=True,
    )

    button_label_map = {
        "检测报告定点返修": "开始检测报告定点返修",
        "中文中转回译（百度翻译）": "开始中文中转回译",
    }
    button_label = button_label_map.get(process_mode, "开始逐段润色")
    if st.button(button_label, use_container_width=True, type="primary"):
        if not uploaded_file:
            st.error("请先上传 Word 文档。")
        elif process_mode == "检测报告定点返修" and not aigc_report_file:
            st.error("请上传 AIGC 检测报告 HTML 或 PDF。")
        elif process_mode == "中文中转回译（百度翻译）" and (not baidu_appid or not baidu_secret_key):
            st.error("请填写百度翻译 APPID 和密钥。")
        elif process_mode != "中文中转回译（百度翻译）" and not api_key.startswith("sk-"):
            st.error("请填写有效的 DeepSeek API Key。")
        else:
            st.session_state.processed_file = None
            st.session_state.tracked_file = None
            progress_bar.progress(0)
            progress_status.caption("任务准备中...")
            if process_mode == "检测报告定点返修":
                st.session_state.output_prefix = "返修版"
                result_bytes = process_report_repair_word(
                    uploaded_file.getvalue(),
                    aigc_report_file.getvalue(),
                    aigc_report_file.name,
                    original_base_file.getvalue() if original_base_file else None,
                    api_key,
                    model_name,
                    concurrency,
                    min_chars,
                    prompt,
                    rewrite_temperature,
                    adaptive_risk_repair,
                    enforce_format_safety,
                    generate_tracked_file,
                    log_container,
                    progress_bar,
                    progress_status,
                )
            elif process_mode == "中文中转回译（百度翻译）":
                st.session_state.output_prefix = "回译版"
                result_bytes = process_word_baidu(
                    uploaded_file.getvalue(),
                    baidu_appid,
                    baidu_secret_key,
                    lang_chain,
                    concurrency,
                    start_paragraph_index,
                    end_paragraph_index,
                    min_chars,
                    enforce_format_safety,
                    generate_tracked_file,
                    log_container,
                    progress_bar,
                    progress_status,
                )
            else:
                st.session_state.output_prefix = "润色版"
                result_bytes = process_word(
                    uploaded_file.getvalue(),
                    api_key,
                    model_name,
                    concurrency,
                    start_paragraph_index,
                    end_paragraph_index,
                    min_chars,
                    prompt,
                    rewrite_temperature,
                    adaptive_risk_repair,
                    enforce_format_safety,
                    generate_tracked_file,
                    log_container,
                    progress_bar,
                    progress_status,
                )

            if result_bytes:
                st.session_state.processed_file = result_bytes

    if st.session_state.processed_file:
        st.success("处理完成，可以下载新文档。")
        st.download_button(
            label="下载处理后的 Word 文档",
            data=st.session_state.processed_file,
            file_name=f"{st.session_state.output_prefix}_{uploaded_file.name if uploaded_file else 'document.docx'}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )
    if st.session_state.tracked_file:
        st.download_button(
            label="下载带修订痕迹的 Word 文档",
            data=st.session_state.tracked_file,
            file_name=f"修订版_{st.session_state.output_prefix}_{uploaded_file.name if uploaded_file else 'document.docx'}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

if st.session_state.rewrite_report:
    st.divider()
    st.subheader("改写对照清单")

    changed_count = sum(1 for item in st.session_state.rewrite_report if item["status"] == "changed")
    unchanged_count = sum(1 for item in st.session_state.rewrite_report if item["status"] == "unchanged")
    failed_count = sum(1 for item in st.session_state.rewrite_report if item["status"] == "failed")
    metric_col_1, metric_col_2, metric_col_3, metric_col_4 = st.columns(4)
    metric_col_1.metric("处理段落", len(st.session_state.rewrite_report))
    metric_col_2.metric("成功替换", changed_count)
    metric_col_3.metric("无变化", unchanged_count)
    metric_col_4.metric("保留原文", failed_count)

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
        status_text_map = {"changed": "已替换", "unchanged": "无变化", "failed": "保留原文"}
        status_text = status_text_map.get(item["status"], item["status"])
        with st.container(border=True):
            st.markdown(
                f'<div class="compare-meta">第 {item["page"]} 页 / 第 {item["paragraph_index"]} 段 - {status_text}</div>',
                unsafe_allow_html=True,
            )
            before_col, after_col = st.columns(2)
            before_col.markdown("**原文**")
            if item["status"] == "changed":
                before_col.markdown(
                    f'<div class="compare-text">{item["old_diff_html"]}</div>',
                    unsafe_allow_html=True,
                )
            else:
                before_col.markdown(
                    f'<div class="compare-text">{html.escape(item["original_text"])}</div>',
                    unsafe_allow_html=True,
                )
            after_col.markdown("**改写后**")
            if item["status"] == "changed":
                after_col.markdown(
                    f'<div class="compare-text">{item["new_diff_html"]}</div>',
                    unsafe_allow_html=True,
                )
            else:
                after_col.markdown(
                    f'<div class="compare-text">{html.escape(item["new_text"])}</div>',
                    unsafe_allow_html=True,
                )
                if item["status"] == "unchanged":
                    after_col.info("无明显文本差异。")
                else:
                    after_col.warning(f"未写回：{item['error']}")
