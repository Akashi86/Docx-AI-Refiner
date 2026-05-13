import pathlib

src = pathlib.Path("app.py").read_text(encoding="utf-8")

# 找到并修复嵌套引号问题
old = (
    'help="不上传则修订版默认比较"当前待返修 Word -> 返修版"。",'
)
new = (
    "help='不上传则修订版默认比较当前待返修 Word -> 返修版。',"
)

if old in src:
    src = src.replace(old, new, 1)
    pathlib.Path("app.py").write_text(src, encoding="utf-8")
    print("fixed")
else:
    # 找到行并打印原始内容
    for i, line in enumerate(src.splitlines(), 1):
        if "待返修 Word" in line:
            print(f"line {i}: {repr(line)}")
