import pathlib

SRC = pathlib.Path("app.py")
text = SRC.read_text(encoding="utf-8")

# 1. Modify suspicious_rewrite_reason
OLD_REASON = """
    if enforce_format_safety and "*" not in original_compact and "*" in new_compact:
        return "markdown_formatting_added"
    if CASUAL_REWRITE_RE.search(new_compact):
        return "casual_rewrite_phrase"
"""

NEW_REASON = """
    if enforce_format_safety and "*" not in original_compact and "*" in new_compact:
        return "markdown_formatting_added"
    if enforce_format_safety and CASUAL_REWRITE_RE.search(new_compact):
        return "casual_rewrite_phrase"
"""
text = text.replace(OLD_REASON.strip('\n'), NEW_REASON.strip('\n'))

# 2. Modify UI help text
OLD_UI = """
    enforce_format_safety = st.checkbox(
        "启用格式安全阀",
        value=True,
        help="开启时，如果输出在原文没有星号的段落里新增 Markdown 星号，会拒绝写回；"
             "关闭后仍保留标题保护、过长扩写等内容安全阀。",
    )
"""

NEW_UI = """
    enforce_format_safety = st.checkbox(
        "启用格式安全阀",
        value=True,
        help="开启时，如果输出在原文没有星号的段落里新增 Markdown 星号，或使用了过于随意的口语化词组（如 don't, can't 等），会拒绝写回；"
             "关闭后将不再拦截这些，但仍保留标题保护、过长扩写等内容安全阀。",
    )
"""
text = text.replace(OLD_UI.strip('\n'), NEW_UI.strip('\n'))

SRC.write_text(text, encoding="utf-8")
print("Patch applied.")
