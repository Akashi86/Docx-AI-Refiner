import pathlib, re

src = pathlib.Path("app.py").read_text(encoding="utf-8")

# 找到所有 help="..." 里面嵌套中文引号的行并打印
count = 0
lines = src.splitlines(keepends=True)
for i, line in enumerate(lines, 1):
    # 中文引号 \u201c \u201d
    if '\u201c' in line or '\u201d' in line:
        print(f"line {i}: {repr(line[:120])}")
        count += 1

print(f"total: {count}")

# 替换所有中文引号为直引号（在 Python 字符串上下文中可能引起问题的才需要处理）
# 实际上最安全的做法是把所有 help="..." 里的中文 " " 改成普通单引号标记
# 用正则找 help="...中文引号..." 模式
def fix_help(m):
    inner = m.group(1)
    inner = inner.replace('\u201c', '').replace('\u201d', '')
    return f'help="{inner}",'

new_src = re.sub(r'help="([^"]*?[\u201c\u201d][^"]*?)",', fix_help, src)

if new_src != src:
    pathlib.Path("app.py").write_text(new_src, encoding="utf-8")
    print("Chinese quotes fixed")
else:
    print("no change")
