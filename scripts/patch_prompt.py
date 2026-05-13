import pathlib

SRC = pathlib.Path("app.py")
text = SRC.read_text(encoding="utf-8")

new_prompt = """    "中国学生 GRE 代入法": \"\"\"你是一个中国的学生，你刚刚通过了gre考试，作文的分数为3，这是你的水平。请你针对以下的英文内容进行重写，想尽力写出大概4分左右水平的英文作文，请记住你是一个中国大学生，可能词汇量和语言习惯和其他国家的学生不同，请你切实的代入后进行写作，减少ai风格的语言习惯，切实代入人类学生，减少模版写作的痕迹，在写作时加入一定的个人色彩。可以在语言上显示出瑕疵，但是你的学术态度端正，请你展现出这种不完美和瑕疵感。注意保留原有的专业术语、数字和引用，不要随意增加或删除客观事实。仅返回重写后的段落文本，不要添加任何Markdown格式（如星号等）。\"\"\",
"""

if "中国学生 GRE 代入法" not in text:
    text = text.replace(
        'PROMPT_TEMPLATES = {\n',
        'PROMPT_TEMPLATES = {\n' + new_prompt,
        1
    )
    SRC.write_text(text, encoding="utf-8")
    print("Prompt added successfully.")
else:
    print("Prompt already exists.")
