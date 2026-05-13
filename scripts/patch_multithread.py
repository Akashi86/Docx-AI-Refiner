import pathlib, re

SRC = pathlib.Path("app.py")
text = SRC.read_text(encoding="utf-8")

# 1. Add call_baidu_chain
CHAIN_FUNC = '''
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
'''
if "def call_baidu_chain" not in text:
    text = text.replace(
        '    return " ".join(item["dst"] for item in data["trans_result"])\n',
        '    return " ".join(item["dst"] for item in data["trans_result"])\n\n' + CHAIN_FUNC,
        1
    )

# 2. Add concurrency to process_word_baidu signature
text = text.replace(
    '    baidu_secret_key,\n    lang_chain,\n',
    '    baidu_secret_key,\n    lang_chain,\n    concurrency,\n',
    1
)

# 3. Replace loop in process_word_baidu
OLD_LOOP = '''
            for idx, task in enumerate(tasks):
                para_no = task["paragraph_index"] + 1
                original_text = task["plain_text"]
                reject_reason = ""
                try:
                    current_text = original_text
                    for step_i in range(len(lang_chain) - 1):
                        from_lang = lang_chain[step_i]
                        to_lang = lang_chain[step_i + 1]
                        add_log(f"第 {para_no} 段 步骤{step_i + 1}/{len(lang_chain) - 1}：{from_lang}→{to_lang}...", "send")
                        render_logs(log_container)
                        current_text = call_baidu_translate(
                            current_text, from_lang, to_lang, baidu_appid, baidu_secret_key
                        )
                        time.sleep(1.1)
                    new_text = current_text.strip()
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

                update_progress(
                    0.06 + (idx + 1) / len(tasks) * 0.84,
                    f"回译中：已完成 {idx + 1}/{len(tasks)} 段。"
                )
                render_logs(log_container)
'''

NEW_LOOP = '''
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
'''
text = text.replace(OLD_LOOP.strip('\n'), NEW_LOOP.strip('\n'))

# 4. Modify UI part
OLD_UI_BAIDU_CONFIG = '''
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
'''

NEW_UI_BAIDU_CONFIG = '''
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
'''
text = text.replace(OLD_UI_BAIDU_CONFIG.strip('\n'), NEW_UI_BAIDU_CONFIG.strip('\n'))

# 5. Modify button click call to pass concurrency
OLD_CALL = '''
            elif process_mode == "中文中转回译（百度翻译）":
                st.session_state.output_prefix = "回译版"
                result_bytes = process_word_baidu(
                    uploaded_file.getvalue(),
                    baidu_appid,
                    baidu_secret_key,
                    lang_chain,
                    start_paragraph_index,
'''
NEW_CALL = '''
            elif process_mode == "中文中转回译（百度翻译）":
                st.session_state.output_prefix = "回译版"
                result_bytes = process_word_baidu(
                    uploaded_file.getvalue(),
                    baidu_appid,
                    baidu_secret_key,
                    lang_chain,
                    concurrency,
                    start_paragraph_index,
'''
text = text.replace(OLD_CALL.strip('\n'), NEW_CALL.strip('\n'))

SRC.write_text(text, encoding="utf-8")
print("patch OK")
