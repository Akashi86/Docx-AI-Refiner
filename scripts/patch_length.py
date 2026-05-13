import pathlib

SRC = pathlib.Path("app.py")
text = SRC.read_text(encoding="utf-8")

OLD_RULES = """
    if not ends_with_sentence_punctuation(original_compact) and ends_with_sentence_punctuation(new_compact):
        if new_words > max(20, int(original_words * 1.6)):
            return "title_or_label_expanded"
    if original_words < 20 and new_words > max(30, int(original_words * 2.5)):
        return "short_text_expanded"
    if original_words >= 30 and new_words > int(original_words * 1.45):
        return "rewrite_too_long"
"""

NEW_RULES = """
    if not ends_with_sentence_punctuation(original_compact) and ends_with_sentence_punctuation(new_compact):
        if new_words > max(20, int(original_words * (2.0 if not enforce_format_safety else 1.6))):
            return "title_or_label_expanded"
    if original_words < 20 and new_words > max(30, int(original_words * (3.5 if not enforce_format_safety else 2.5))):
        return "short_text_expanded"
    if original_words >= 30 and new_words > int(original_words * (2.5 if not enforce_format_safety else 1.45)):
        return "rewrite_too_long"
"""

text = text.replace(OLD_RULES.strip('\n'), NEW_RULES.strip('\n'))

OLD_UI = """
    enforce_format_safety = st.checkbox(
        "启用格式安全阀",
        value=True,
        help="开启时，如果输出在原文没有星号的段落里新增 Markdown 星号，或使用了过于随意的口语化词组（如 don't, can't 等），会拒绝写回；"
             "关闭后将不再拦截这些，但仍保留标题保护、过长扩写等内容安全阀。",
    )
"""

NEW_UI = """
    enforce_format_safety = st.checkbox(
        "启用格式安全阀",
        value=True,
        help="开启时，禁止输出新增 Markdown 星号或过于随意的口语化词组，并严格限制段落扩写长度（最大 1.45 倍）；"
             "关闭后不仅允许星号和口语化，还会将字数扩写限制大幅放宽（最高允许 2.5 倍），以适应大量废话扩写。",
    )
"""
text = text.replace(OLD_UI.strip('\n'), NEW_UI.strip('\n'))

SRC.write_text(text, encoding="utf-8")
print("Patch applied for length limits.")
