# AGENTS.md

## 项目规则

- git 提交信息生成用中文。
- 默认工作目录是 `E:\work\Docx-AI-Refiner`。
- 当前主分支是 `main`，远端是 `git@github.com:dezhengliu490-del/Docx-AI-Refiner.git`。

## 项目目标

这是一个 Streamlit 应用，用于处理用户上传的 Word 论文 `.docx`：

1. 读取 Word 文档正文段落。
2. 将正文逐段发送给 DeepSeek API，用指定提示词改写。
3. 将改写结果写回 Word。
4. 尽量不改动原有 Word 格式。
5. 在网页上展示原文和改写后的对照清单，并支持下载清单。

根本需求是：**AI 改正文，但不要破坏 Word 论文原有格式**。

## 当前实现进度

### 已完成

- 已把项目初始化为 Git 仓库并推送到 GitHub。
- 已安装 GitHub CLI，但当前系统里 `gh` 需要新终端刷新 PATH 后才能直接使用。
- 已将原先“删除 run 后重建 run”的实现，改为 **XML 文本节点替换**。
- 当前写回逻辑不会重建段落、run、样式节点，而是只替换原有 `w:t` 文本节点。
- 已增加 DeepSeek API 调用的有限重试，避免无限递归卡死。
- 已修复日志 HTML 注入风险：日志内容会先做 `html.escape()`。
- 已新增 `.gitignore`，忽略 `__pycache__/`、`.env`、`.streamlit/secrets.toml` 等。
- 默认提示词已改成英文：

```text
Without changing the paragraph structure or the order of information, rewrite the text below to reduce Al vibes. Do not start sentences with generic stock phrases ("in conclusion," "moreover," etc.); use more specific, context-tied transitions instead. Make each paragraph's first sentence feel like a natural continuation of the previous one. Favor concrete verbs over stacks of abstract nouns, keep terminology precise, and split long sentences into 2- 3 shorter ones.
```

- 已新增“改写对照清单”：
  - 直接平铺展示，不使用下拉。
  - 左侧显示原文，右侧显示改写后。
  - 原文和改写后都有差异高亮。
  - 原文栏中，被删除内容用红色删除线，被替换的原内容用红色高亮。
  - 改写后栏中，新增内容用绿色高亮，替换后的新内容用黄色高亮。
  - 支持下载 `rewrite_report.json` 和 `rewrite_report.csv`。
- 已增加非正文过滤：
  - 跳过 `Title`、`Heading`、`TOC`、`Caption`、`Header`、`Footer`、`Reference` 等样式。
  - 跳过短标题形态，如 `Abstract`、`Chapter`、`References`、`Appendix` 等。
- 已增加状态区分：
  - `changed`：成功替换。
  - `unchanged`：AI 返回内容与原文一致。
  - `failed`：AI 返回不合规或调用失败，保留原文。

## 当前关键实现

主要代码在 `app.py`。

核心流程：

1. `.docx` 按 zip 读取。
2. 读取 `word/document.xml`。
3. 扫描 `w:p` 段落。
4. 对每段收集所有 `w:t` 文本节点。
5. 发送给 AI 前，把每个文本节点包装为：

```xml
<seg id="0">...</seg><seg id="1">...</seg>
```

6. 系统提示词要求 AI 必须保留所有 `seg id`。
7. AI 返回后，程序解析所有 `seg`。
8. 如果返回片段数量、id、顺序不合规，则该段不写回，保留原文。
9. 如果合规，则把每个 `seg` 内容写回对应的 `w:t.text`。
10. 打包生成新的 `.docx`。

## 已知问题

### 1. 多 `seg` 机制安全但成功率不够高

当前方案为了最大限度保护 Word 格式，要求 AI 返回和原文文本节点数量完全一致。

问题是 Word 经常会把一个普通段落拆成很多碎 `w:t` 节点。例如一段正文可能有 16 个文本节点，AI 只返回 15 个片段时，程序会报：

```text
AI 返回了 15 个片段，但期望 16 个。
```

这时程序会把该段标记为“保留原文”。这是安全行为，避免错位写回导致格式损坏。

### 2. 页码判断不一定精确

当前页码依赖 `w:lastRenderedPageBreak` 和手动分页符。这是 Word 的渲染缓存，不是可靠分页引擎。

如果文档没有保存过分页缓存，或不同机器字体/页面设置不同，页码可能不准。

### 3. 正文过滤仍需根据真实论文继续调优

当前已跳过常见标题和目录样式，但不同 Word 模板的样式名可能不同。

如果用户文档里标题样式不是 `Heading`，或者正文不是 `Normal/Body Text`，需要继续调整 `is_body_paragraph()`。

### 4. 差异高亮是基于空格分词

当前 `difflib` 用 `text.split()` 做词级对比。英文论文基本可用。

中文或中英混排时，高亮粒度可能不够细，需要改成字符级或更智能的 tokenizer。

## 下一步建议

### 优先级最高：提高 AI 写回成功率

建议把当前“多 seg 精确映射”改成更稳的双模式：

1. **保守模式，当前实现**：
   - 每个 `w:t` 一个 `seg`。
   - 格式保护最强。
   - AI 少返回一个片段就保留原文。

2. **高成功率模式，推荐新增**：
   - 每个段落只给 AI 一个完整文本块。
   - AI 只返回改写后的纯文本。
   - 程序负责写回。
   - 写回策略可选：
     - 把改写全文写入第一个 `w:t`，其余 `w:t` 清空。
     - 或按原文本节点长度比例，把新文本分配回多个 `w:t`。

推荐先做“高成功率模式”：**改写全文写入第一个 `w:t`，其余 `w:t` 清空**。

这个方案仍然保留段落、run 和样式结构，不重建 XML，但会让段内局部 run 样式粒度下降。如果论文正文大多数是统一格式，这个折中很实用。

### 建议新增 UI 开关

在页面上加一个写回模式选择：

- `严格保格式模式`：当前多 `seg` 模式。
- `高成功率模式`：整段改写，写入首个文本节点。

默认建议使用 `高成功率模式`，因为用户更关心正文确实被改写。

### 建议继续优化清单

- 给每段增加“复制原文/复制改写后”按钮。
- 对失败段落提供“重试该段”按钮。
- 对 `unchanged` 段落提供“加强改写提示词后重试”按钮。
- 支持只下载 changed 段落的对照清单。

## 开发和验证命令

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

运行应用：

```powershell
python -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501
```

编译检查：

```powershell
python -c "import pathlib; compile(pathlib.Path('app.py').read_text(encoding='utf-8'), 'app.py', 'exec'); print('COMPILE OK')"
```

查看状态：

```powershell
git status --short --branch
```

推送：

```powershell
git add app.py AGENTS.md .gitignore requirements.txt
git commit -m "中文提交信息"
git push
```

## 重要经验

- 不要再使用“删除旧 run，再创建新 run”的方式写回 Word，这会破坏字体、字号、颜色、链接、脚注、域、批注、编号等结构。
- 操作 `.docx` 时应尽量保留原 XML 树，只改 `w:t.text`。
- AI 不可靠，凡是让 AI 保留结构化标签的地方，都必须做严格校验。
- 校验失败时宁愿保留原文，也不要强行写回。
- 对照清单应该由程序根据原文和新文生成，不应该完全依赖 AI 自己描述“改了什么”。
- 当前 PowerShell 输出中文可能显示乱码，但文件本身是 UTF-8，修改时注意不要误判为文件损坏。
