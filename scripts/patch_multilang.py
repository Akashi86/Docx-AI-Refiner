"""
把百度翻译回译改为多语言链条模式
1. process_word_baidu() 加 lang_chain 参数，两步翻译改为循环
2. UI 加语言链条选择器（预设 + 中间语言 multiselect）
3. 按钮调用传 lang_chain
"""
import pathlib

SRC = pathlib.Path("app.py")
text = SRC.read_text(encoding="utf-8")

# ── 1. process_word_baidu 签名加 lang_chain ────────────────────────────────
OLD_SIG = (
    "def process_word_baidu(\n"
    "    file_bytes,\n"
    "    baidu_appid,\n"
    "    baidu_secret_key,\n"
    "    start_paragraph_index,\n"
    "    end_paragraph_index,\n"
    "    min_chars,\n"
    "    enforce_format_safety,\n"
    "    generate_tracked_file,\n"
    "    log_container,\n"
    "    progress_bar,\n"
    "    progress_status=None,\n"
    "):"
)
NEW_SIG = (
    "def process_word_baidu(\n"
    "    file_bytes,\n"
    "    baidu_appid,\n"
    "    baidu_secret_key,\n"
    "    lang_chain,\n"
    "    start_paragraph_index,\n"
    "    end_paragraph_index,\n"
    "    min_chars,\n"
    "    enforce_format_safety,\n"
    "    generate_tracked_file,\n"
    "    log_container,\n"
    "    progress_bar,\n"
    "    progress_status=None,\n"
    "):"
)
assert OLD_SIG in text, "process_word_baidu sig not found"
text = text.replace(OLD_SIG, NEW_SIG, 1)

# ── 2. 日志里"英→中→英"改为动态链条显示 ──────────────────────────────────
OLD_LOG = (
    '            add_log(f"共找到 {len(tasks)} 个待处理段落，开始逐段回译（英→中→英）。", "info")\n'
    '            add_log("注意：百度翻译免费版 QPS=1，每段约需 2~3 秒，请耐心等待。", "warn")\n'
    '            update_progress(0.06, f"已找到 {len(tasks)} 个待处理段落，开始回译...")'
)
NEW_LOG = (
    '            chain_label = " → ".join(lang_chain)\n'
    '            steps = len(lang_chain) - 1\n'
    '            add_log(f"共找到 {len(tasks)} 个待处理段落，翻译链条：{chain_label}（共 {steps} 步）。", "info")\n'
    '            add_log(f"注意：百度翻译免费版 QPS=1，每段约需 {steps * 1.1:.0f}~{steps * 2:.0f} 秒，请耐心等待。", "warn")\n'
    '            update_progress(0.06, f"已找到 {len(tasks)} 个待处理段落，开始回译...")'
)
assert OLD_LOG in text, "log block not found"
text = text.replace(OLD_LOG, NEW_LOG, 1)

# ── 3. 替换两步翻译为循环 ───────────────────────────────────────────────────
OLD_TRANS = (
    '                try:\n'
    '                    add_log(f"第 {para_no} 段，第 {task[\'page\']} 页：英→中...", "send")\n'
    '                    render_logs(log_container)\n'
    '                    chinese_text = call_baidu_translate(\n'
    '                        original_text, "en", "zh", baidu_appid, baidu_secret_key\n'
    '                    )\n'
    '                    time.sleep(1.1)\n'
    '                    add_log(f"第 {para_no} 段：中→英...", "send")\n'
    '                    render_logs(log_container)\n'
    '                    english_text = call_baidu_translate(\n'
    '                        chinese_text, "zh", "en", baidu_appid, baidu_secret_key\n'
    '                    )\n'
    '                    time.sleep(1.1)\n'
    '                    new_text = english_text.strip()'
)
NEW_TRANS = (
    '                try:\n'
    '                    current_text = original_text\n'
    '                    for step_i in range(len(lang_chain) - 1):\n'
    '                        from_lang = lang_chain[step_i]\n'
    '                        to_lang = lang_chain[step_i + 1]\n'
    '                        add_log(f"第 {para_no} 段 步骤{step_i + 1}/{len(lang_chain) - 1}：{from_lang}→{to_lang}...", "send")\n'
    '                        render_logs(log_container)\n'
    '                        current_text = call_baidu_translate(\n'
    '                            current_text, from_lang, to_lang, baidu_appid, baidu_secret_key\n'
    '                        )\n'
    '                        time.sleep(1.1)\n'
    '                    new_text = current_text.strip()'
)
assert OLD_TRANS in text, "two-step translation block not found"
text = text.replace(OLD_TRANS, NEW_TRANS, 1)

# ── 4. rewrite_rounds 改为实际步数 ─────────────────────────────────────────
OLD_ROUNDS = '                        "rewrite_rounds": 2,'
NEW_ROUNDS = '                        "rewrite_rounds": len(lang_chain) - 1,'
assert OLD_ROUNDS in text, "rewrite_rounds line not found"
text = text.replace(OLD_ROUNDS, NEW_ROUNDS, 1)

# ── 5. UI：把"回译说明"区域替换为链条选择器 ─────────────────────────────────
OLD_UI = (
    '    else:\n'
    '        prompt = ""\n'
    '        rewrite_temperature = 0.55\n'
    '        adaptive_risk_repair = False\n'
    '        st.subheader("5. 回译说明")\n'
    '        st.info(\n'
    '            "回译模式不需要提示词和改写强度。\\n\\n"\n'
    '            "流程：英文段落 → 百度翻译成中文 → 再翻译回英文。\\n"\n'
    '            "所有已有安全阀（长度控制、格式保护、protected text 检测等）均正常生效。"\n'
    '        )'
)
NEW_UI = (
    '    else:\n'
    '        prompt = ""\n'
    '        rewrite_temperature = 0.55\n'
    '        adaptive_risk_repair = False\n'
    '        st.subheader("5. 多语言回译链条")\n'
    '        BAIDU_LANG_OPTIONS = {\n'
    '            "英语 (en)": "en",\n'
    '            "中文 (zh)": "zh",\n'
    '            "法语 (fra)": "fra",\n'
    '            "俄语 (ru)": "ru",\n'
    '            "日语 (jp)": "jp",\n'
    '            "德语 (de)": "de",\n'
    '            "西班牙语 (spa)": "spa",\n'
    '            "阿拉伯语 (ara)": "ara",\n'
    '            "韩语 (kor)": "kor",\n'
    '            "葡萄牙语 (pt)": "pt",\n'
    '        }\n'
    '        PRESET_CHAINS = {\n'
    '            "英→中→英（2步，最快）": ["en", "zh", "en"],\n'
    '            "英→法→中→英（3步）": ["en", "fra", "zh", "en"],\n'
    '            "英→法→中→俄→英（4步）": ["en", "fra", "zh", "ru", "en"],\n'
    '            "英→日→法→中→英（4步）": ["en", "jp", "fra", "zh", "en"],\n'
    '            "英→德→日→中→俄→英（5步）": ["en", "de", "jp", "zh", "ru", "en"],\n'
    '            "自定义": None,\n'
    '        }\n'
    '        preset_name = st.selectbox(\n'
    '            "预设链条",\n'
    '            list(PRESET_CHAINS.keys()),\n'
    '            index=1,\n'
    '            help="选择预设翻译链条，步数越多混淆效果越强，但耗时更长。",\n'
    '        )\n'
    '        if PRESET_CHAINS[preset_name] is not None:\n'
    '            lang_chain = PRESET_CHAINS[preset_name]\n'
    '            chain_display = " → ".join(\n'
    '                next(k for k, v in BAIDU_LANG_OPTIONS.items() if v == c)\n'
    '                for c in lang_chain\n'
    '            )\n'
    '            st.caption(f"当前链条：{chain_display}")\n'
    '        else:\n'
    '            st.caption("自定义中间语言（起点英语 en，终点英语 en 固定，选择中间经过的语言）")\n'
    '            mid_langs = st.multiselect(\n'
    '                "中间语言（按顺序选择）",\n'
    '                [k for k in BAIDU_LANG_OPTIONS if k != "英语 (en)"],\n'
    '                default=["法语 (fra)", "中文 (zh)"],\n'
    '                key="custom_mid_langs",\n'
    '            )\n'
    '            if not mid_langs:\n'
    '                st.warning("请至少选择一种中间语言。")\n'
    '                mid_langs = ["中文 (zh)"]\n'
    '            lang_chain = ["en"] + [BAIDU_LANG_OPTIONS[k] for k in mid_langs] + ["en"]\n'
    '            chain_display = " → ".join(\n'
    '                next(k for k, v in BAIDU_LANG_OPTIONS.items() if v == c)\n'
    '                for c in lang_chain\n'
    '            )\n'
    '            st.caption(f"当前链条：{chain_display}")\n'
    '        steps = len(lang_chain) - 1\n'
    '        st.info(\n'
    '            f"翻译步数：{steps} 步，每段约需 {steps * 1.1:.0f}~{steps * 2:.0f} 秒（免费版 QPS=1）。\\n"\n'
    '            "所有格式安全阀、长度保护、protected text 检测均正常生效。"\n'
    '        )'
)
assert OLD_UI in text, "UI else block not found"
text = text.replace(OLD_UI, NEW_UI, 1)

# ── 6. 按钮调用处加 lang_chain 参数 ────────────────────────────────────────
OLD_CALL = (
    '            elif process_mode == "中文中转回译（百度翻译）":\n'
    '                st.session_state.output_prefix = "回译版"\n'
    '                result_bytes = process_word_baidu(\n'
    '                    uploaded_file.getvalue(),\n'
    '                    baidu_appid,\n'
    '                    baidu_secret_key,\n'
    '                    start_paragraph_index,\n'
    '                    end_paragraph_index,\n'
    '                    min_chars,\n'
    '                    enforce_format_safety,\n'
    '                    generate_tracked_file,\n'
    '                    log_container,\n'
    '                    progress_bar,\n'
    '                    progress_status,\n'
    '                )'
)
NEW_CALL = (
    '            elif process_mode == "中文中转回译（百度翻译）":\n'
    '                st.session_state.output_prefix = "回译版"\n'
    '                result_bytes = process_word_baidu(\n'
    '                    uploaded_file.getvalue(),\n'
    '                    baidu_appid,\n'
    '                    baidu_secret_key,\n'
    '                    lang_chain,\n'
    '                    start_paragraph_index,\n'
    '                    end_paragraph_index,\n'
    '                    min_chars,\n'
    '                    enforce_format_safety,\n'
    '                    generate_tracked_file,\n'
    '                    log_container,\n'
    '                    progress_bar,\n'
    '                    progress_status,\n'
    '                )'
)
assert OLD_CALL in text, "button call block not found"
text = text.replace(OLD_CALL, NEW_CALL, 1)

SRC.write_text(text, encoding="utf-8")
print("patch OK")
