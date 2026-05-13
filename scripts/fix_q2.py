import pathlib, re

src = pathlib.Path("app.py").read_text(encoding="utf-8")

# 找到所有包含中文弯引号的 help= 行并修复
# 思路：在 help="..." 内部，把嵌套的 " 改为空（或用书名号替换）
fixes = [
    (
        'help="强度越高，句式和措辞变化越大；如果检测结果仍偏高，优先尝试"最大强改写"。",',
        'help="强度越高，句式和措辞变化越大；如果检测结果仍偏高，优先尝试【最大强改写】。",',
    ),
]

for old, new in fixes:
    if old in src:
        src = src.replace(old, new, 1)
        print(f"fixed: {old[:40]}...")
    else:
        # 尝试用行号方式打印
        for i, line in enumerate(src.splitlines(), 1):
            if "最大强改写" in line and "help=" in line:
                print(f"line {i}: {repr(line)}")

pathlib.Path("app.py").write_text(src, encoding="utf-8")
print("done")
