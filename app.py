import streamlit as st
import zipfile
import io
import xml.etree.ElementTree as ET
import re
import requests
import time
import concurrent.futures

# ================= 命名空间注册 (核心：防止Word文件损坏) =================
# 在 Python 中处理 Word 的 XML 时，必须注册这些命名空间，否则保存后 Word 会报错
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

W_NS = f"{{{NAMESPACES['w']}}}"

# ================= 页面配置 =================
st.set_page_config(page_title="AI 学术论文精修工具", page_icon="📄", layout="wide")

st.markdown("""
    <style>
    .log-box {
        background-color: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 10px;
        font-family: monospace;
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
    </style>
""", unsafe_allow_html=True)

# 初始化 Session State
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'processed_file' not in st.session_state:
    st.session_state.processed_file = None

def add_log(msg, type="normal"):
    """添加日志的辅助函数"""
    time_str = time.strftime("%H:%M:%S")
    if type == "success": html_msg = f'<div class="log-entry log-success">[{time_str}] {msg}</div>'
    elif type == "send": html_msg = f'<div class="log-entry log-send">[{time_str}] {msg}</div>'
    elif type == "warn": html_msg = f'<div class="log-entry log-warn">[{time_str}] {msg}</div>'
    elif type == "err": html_msg = f'<div class="log-entry log-err">[{time_str}] {msg}</div>'
    elif type == "info": html_msg = f'<div class="log-entry log-info">[{time_str}] {msg}</div>'
    else: html_msg = f'<div class="log-entry">[{time_str}] {msg}</div>'
    
    st.session_state.logs.append(html_msg)

def call_deepseek(text, system_prompt, api_key, model_name, task_id, attempt=1):
    """调用 DeepSeek API (带重试机制，绝不轻易熔断)"""
    url = "https://api.deepseek.com/chat/completions"
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}
        ],
        "temperature": 0.3
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    try:
        response = requests.post(url, json=payload, headers=headers, timeout=120) # 设置较长超时应对R1
        if response.status_code != 200:
            error_msg = response.text
            if response.status_code == 401: raise Exception("AuthError: API Key 错误或无效。")
            if response.status_code == 402: raise Exception("BalanceError: 账号余额不足。")
            if response.status_code in [400, 404]: raise Exception(f"ModelError: 模型 [{model_name}] 错误。")
            raise Exception(f"HTTP {response.status_code}: {error_msg}")
        
        data = response.json()
        return data['choices'][0]['message']['content'].strip()

    except Exception as e:
        is_fatal = any(err in str(e) for err in ["AuthError", "BalanceError", "ModelError"])
        if is_fatal:
            raise e
        
        delay = min(attempt * 4, 30)
        time.sleep(delay)
        return call_deepseek(text, system_prompt, api_key, model_name, task_id, attempt + 1)

def process_word(file_bytes, api_key, model_name, concurrency, start_page, end_page, min_words, prompt, log_container, progress_bar):
    """核心处理逻辑"""
    st.session_state.logs = []
    
    try:
        add_log("📥 正在解压并解析 Word 文档...", "normal")
        log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)
        
        # 1. 在内存中操作 zip (docx)
        doc_zip = zipfile.ZipFile(io.BytesIO(file_bytes), 'r')
        xml_content = doc_zip.read('word/document.xml')
        tree = ET.ElementTree(ET.fromstring(xml_content))
        root = tree.getroot()

        # 2. 扫描并构建任务队列
        tasks = []
        current_page = 1
        end_page_num = end_page if end_page else float('inf')
        
        paragraphs = list(root.iter(f'{W_NS}p'))
        add_log("🔍 正在扫描页码并构建任务队列...", "normal")

        for idx, p in enumerate(paragraphs):
            # 计算分页
            page_breaks = len(p.findall(f'.//{W_NS}lastRenderedPageBreak'))
            for br in p.findall(f'.//{W_NS}br'):
                if br.get(f'{W_NS}type') == 'page':
                    page_breaks += 1
            current_page += page_breaks

            if current_page < start_page or current_page > end_page_num:
                continue

            # 提取文本与格式
            tagged_text = ""
            runs = list(p.findall(f'{W_NS}r'))
            for r in runs:
                t_elem = r.find(f'{W_NS}t')
                if t_elem is None or not t_elem.text:
                    continue
                text = t_elem.text

                rPr = r.find(f'{W_NS}rPr')
                is_bold = rPr is not None and rPr.find(f'{W_NS}b') is not None
                is_italic = rPr is not None and rPr.find(f'{W_NS}i') is not None
                is_underline = rPr is not None and rPr.find(f'{W_NS}u') is not None

                if is_bold: text = f"<b>{text}</b>"
                elif is_italic: text = f"<i>{text}</i>"
                elif is_underline: text = f"<u>{text}</u>"
                
                tagged_text += text

            plain_text = re.sub(r'<[^>]+>', '', tagged_text).strip()
            if len(plain_text) >= min_words:
                tasks.append({
                    "original_index": idx,
                    "page": current_page,
                    "tagged_text": tagged_text,
                    "p_node": p,
                    "runs": runs
                })

        if not tasks:
            add_log("⚠️ 在设定范围内没有找到符合字数要求的段落。", "warn")
            log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)
            return None

        add_log(f"🎯 构建完毕！共提取出 {len(tasks)} 个段落，开始 {concurrency} 并发处理...", "info")
        log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)

        # 3. 并发处理
        completed_tasks = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=concurrency) as executor:
            future_to_task = {}
            for task in tasks:
                future = executor.submit(call_deepseek, task['tagged_text'], prompt, api_key, model_name, task['original_index'])
                future_to_task[future] = task
                add_log(f"[第 {task['original_index']+1} 段] (第{task['page']}页) 🚀 已发往模型...", "send")
                log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)

            for future in concurrent.futures.as_completed(future_to_task):
                task = future_to_task[future]
                try:
                    new_tagged_text = future.result()
                    add_log(f"[第 {task['original_index']+1} 段] ✅ 润色完成！", "success")
                    task['new_tagged_text'] = new_tagged_text
                except Exception as exc:
                    add_log(f"❌ [第 {task['original_index']+1} 段] 发生致命错误: {exc}", "err")
                    task['new_tagged_text'] = task['tagged_text'] # 失败则保留原文
                
                completed_tasks += 1
                progress_bar.progress(completed_tasks / len(tasks))
                log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)

        # 4. 将结果写回 XML
        add_log("🎉 队列处理完毕，正在重新植入格式与封存文档...", "info")
        log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)
        
        for task in tasks:
            p = task['p_node']
            # 清理旧 run
            for r in task['runs']:
                p.remove(r)
            
            # 解析并植入新 run
            parts = [pt for pt in re.split(r'(<b>.*?</b>|<i>.*?</i>|<u>.*?</u>)', task['new_tagged_text']) if pt]
            for part in parts:
                new_r = ET.Element(f'{W_NS}r')
                has_format = False
                rPr = ET.Element(f'{W_NS}rPr')

                if part.startswith('<b>') and part.endswith('</b>'):
                    part = part[3:-4]
                    ET.SubElement(rPr, f'{W_NS}b')
                    has_format = True
                elif part.startswith('<i>') and part.endswith('</i>'):
                    part = part[3:-4]
                    ET.SubElement(rPr, f'{W_NS}i')
                    has_format = True
                elif part.startswith('<u>') and part.endswith('</u>'):
                    part = part[3:-4]
                    u = ET.SubElement(rPr, f'{W_NS}u')
                    u.set(f'{W_NS}val', 'single')
                    has_format = True

                if has_format:
                    new_r.append(rPr)
                
                new_t = ET.SubElement(new_r, f'{W_NS}t')
                new_t.set('xml:space', 'preserve')
                new_t.text = part
                p.append(new_r)

        # 5. 重新打包 zip (docx)
        output_io = io.BytesIO()
        with zipfile.ZipFile(output_io, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            # 复制原文档中未修改的文件
            for item in doc_zip.infolist():
                if item.filename != 'word/document.xml':
                    out_zip.writestr(item, doc_zip.read(item.filename))
            # 写入修改后的 document.xml
            xml_str = ET.tostring(root, encoding='utf-8', xml_declaration=True).decode('utf-8')
            out_zip.writestr('word/document.xml', xml_str)

        add_log("✅ 新文档打包完成，请点击下方按钮下载！", "success")
        log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)
        
        return output_io.getvalue()

    except Exception as e:
        add_log(f"❌ 核心流程发生异常: {str(e)}", "err")
        log_container.markdown(f'<div class="log-box">{"".join(st.session_state.logs)}</div>', unsafe_allow_html=True)
        return None

# ================= UI 布局 =================

st.title("⚡ AI 论文并发精修工具 (Streamlit 专属版)")
st.markdown("纯 Python 底层驱动，原生突破沙箱下载限制，支持高并发处理，极其适合部署到公网。")

col_left, col_right = st.columns([1, 1.2])

with col_left:
    st.subheader("⚙️ 1. 核心与范围配置")
    api_key = st.text_input("DeepSeek API Key (sk-开头)", type="password")
    
    col_1, col_2 = st.columns(2)
    model_name = col_1.selectbox("模型选择", ["deepseek-chat", "deepseek-reasoner"])
    concurrency = col_2.slider("并发请求数", min_value=1, max_value=10, value=3, help="建议 3-5，避免触发 API 限流")
    
    col_3, col_4 = st.columns(2)
    start_page = col_3.number_input("开始页", min_value=1, value=1)
    end_page_input = col_4.number_input("结束页 (0代表至末尾)", min_value=0, value=0)
    end_page = int(end_page_input) if end_page_input > 0 else None
    
    min_words = st.slider("过滤小标题 (< N 字)", min_value=5, max_value=50, value=15, step=5)
    
    st.subheader("🧠 2. 设定提示词")
    default_prompt = """Without changing the paragraph structure or the order of information, rewrite the text below to reduce Al vibes. Do not start sentences with generic stock phrases ("in conclusion," "moreover," etc.); use more specific, context-tied transitions instead. Make each paragraph's first sentence feel like a natural continuation of the previous one. Favor concrete verbs over stacks of abstract nouns, keep terminology precise, and split long sentences into 2- 3 shorter ones.

【CRITICAL RULE FOR FORMATTING】: 
The original text contains HTML tags like <b>...</b>(bold), <i>...</i>(italic), or <u>...</u>(underline) to mark important terms. You MUST preserve these exact HTML tags and wrap them around the corresponding rewritten terms! Do NOT lose any tags. Return ONLY the rewritten text without any additional explanations."""
    prompt = st.text_area("System Prompt", value=default_prompt, height=200)

    st.subheader("📄 3. 上传文档")
    uploaded_file = st.file_uploader("选择 Word 文档 (.docx)", type=["docx"])

with col_right:
    st.subheader("🖥️ 运行监控面板")
    
    # 进度条和日志容器
    progress_bar = st.progress(0)
    log_container = st.empty()
    
    # 初始化渲染空日志框
    log_container.markdown(f'<div class="log-box"><div class="log-entry">等待任务开始...</div></div>', unsafe_allow_html=True)
    
    # 执行按钮
    if st.button("🚀 开始执行精修", use_container_width=True, type="primary"):
        if not uploaded_file:
            st.error("请先上传 Word 文档！")
        elif not api_key.startswith("sk-"):
            st.error("请填写正确的 API Key！")
        else:
            file_bytes = uploaded_file.read()
            result_bytes = process_word(
                file_bytes, api_key, model_name, concurrency, 
                start_page, end_page, min_words, prompt, 
                log_container, progress_bar
            )
            
            if result_bytes:
                st.session_state.processed_file = result_bytes

    # Streamlit 原生的完美下载按钮
    if st.session_state.processed_file:
        st.success("🎉 全部处理完毕，格式已无损缝合！")
        st.download_button(
            label="📥 下载最终定稿 (突破所有下载限制)",
            data=st.session_state.processed_file,
            file_name=f"Streamlit定稿_{uploaded_file.name if uploaded_file else 'document.docx'}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary"
        )