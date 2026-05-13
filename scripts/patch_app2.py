"""
重构：把中文中转回译从独立Tab改为处理模式之一，完整继承主流程安全措施。
1. 修复 call_baidu_translate 的 \\n join bug -> 改为空格
2. 删除 back_translate_docx_baidu 函数
3. 新增 process_word_baidu 函数（完整继承主流程框架）
4. 删除 bt_result_file / bt_report session_state 初始化
5. 重写 UI 部分（从 st.title 到文件末尾）：
   - 去掉 Tab，恢复 col 布局
   - process_mode 增加"中文中转回译（百度翻译）"选项
   - 条件显示 DeepSeek / 百度 配置
   - 条件显示提示词/改写强度/风险扫描设置
   - 按钮分三路调用
"""
import pathlib, re

SRC = pathlib.Path("app.py")
text = SRC.read_text(encoding="utf-8")

# ── 1. 修复 \n join bug ────────────────────────────────────────────────────
OLD_JOIN = '    return "\\n".join(item["dst"] for item in data["trans_result"])'
NEW_JOIN = '    return " ".join(item["dst"] for item in data["trans_result"])'
assert OLD_JOIN in text, "join bug line not found"
text = text.replace(OLD_JOIN, NEW_JOIN, 1)

# ── 2. 删除 back_translate_docx_baidu 函数 ─────────────────────────────────
text = re.sub(
    r'\ndef back_translate_docx_baidu\(.*?\n(?=\ndef |\nst\.)',
    '\n',
    text, count=1, flags=re.DOTALL
)

# ── 3. 删除 bt session_state 初始化 ────────────────────────────────────────
text = text.replace(
    'if "bt_result_file" not in st.session_state:\n'
    '    st.session_state.bt_result_file = None\n'
    'if "bt_report" not in st.session_state:\n'
    '    st.session_state.bt_report = []\n',
    ''
)

# ── 4. 新增 process_word_baidu 函数 ────────────────────────────────────────
PROCESS_BAIDU = '''
def process_word_baidu(
    file_bytes,
    baidu_appid,
    baidu_secret_key,
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

            add_log(f"共找到 {len(tasks)} 个待处理段落，开始逐段回译（英→中→英）。", "info")
            add_log("注意：百度翻译免费版 QPS=1，每段约需 2~3 秒，请耐心等待。", "warn")
            update_progress(0.06, f"已找到 {len(tasks)} 个待处理段落，开始回译...")
            render_logs(log_container)

            for idx, task in enumerate(tasks):
                para_no = task["paragraph_index"] + 1
                original_text = task["plain_text"]
                reject_reason = ""
                try:
                    add_log(f"第 {para_no} 段，第 {task['page']} 页：英→中...", "send")
                    render_logs(log_container)
                    chinese_text = call_baidu_translate(
                        original_text, "en", "zh", baidu_appid, baidu_secret_key
                    )
                    time.sleep(1.1)
                    add_log(f"第 {para_no} 段：中→英...", "send")
                    render_logs(log_container)
                    english_text = call_baidu_translate(
                        chinese_text, "zh", "en", baidu_appid, baidu_secret_key
                    )
                    time.sleep(1.1)
                    new_text = english_text.strip()
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
                        "rewrite_rounds": 2,
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

                update_progress(
                    0.06 + (idx + 1) / len(tasks) * 0.84,
                    f"回译中：已完成 {idx + 1}/{len(tasks)} 段。"
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

'''

# 插到 process_report_repair_word 定义之前
ANCHOR = '\ndef process_report_repair_word('
assert ANCHOR in text, "process_report_repair_word anchor not found"
text = text.replace(ANCHOR, PROCESS_BAIDU + '\ndef process_report_repair_word(', 1)

# ── 5. 重写 UI 区（从 st.title 到文件末尾） ────────────────────────────────
UI_START_MARKER = '\nst.title("AI Word 论文逐段润色工具")'
assert UI_START_MARKER in text, "st.title marker not found"
cut_pos = text.index(UI_START_MARKER)
text = text[:cut_pos]

NEW_UI = r'''
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
        concurrency = 1
        baidu_appid = st.text_input("百度翻译 APPID", key="baidu_appid")
        baidu_secret_key = st.text_input("百度翻译密钥（Secret Key）", type="password", key="baidu_secret_key")
        st.caption(
            "在 [百度翻译开放平台](https://fanyi-api.baidu.com/) 免费注册，"
            "每月5万字符免费额度，QPS=1，每段约需 2~3 秒。"
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
                help="不上传则修订版默认比较"当前待返修 Word -> 返修版"。",
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
            help="强度越高，句式和措辞变化越大；如果检测结果仍偏高，优先尝试"最大强改写"。",
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
        st.subheader("5. 回译说明")
        st.info(
            "回译模式不需要提示词和改写强度。\n\n"
            "流程：英文段落 → 百度翻译成中文 → 再翻译回英文。\n"
            "所有已有安全阀（长度控制、格式保护、protected text 检测等）均正常生效。"
        )

    enforce_format_safety = st.checkbox(
        "启用格式安全阀",
        value=True,
        help="开启时，如果输出在原文没有星号的段落里新增 Markdown 星号，会拒绝写回；"
             "关闭后仍保留标题保护、过长扩写等内容安全阀。",
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
'''

text = text.rstrip() + NEW_UI

SRC.write_text(text, encoding="utf-8")
print("patch OK")
