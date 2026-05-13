# AGENTS.md

## 项目规则

- git 提交信息生成用中文。
- 默认工作目录是 `E:\work\Docx-AI-Refiner`。
- 当前主分支是 `main`。
- 远端仓库目前推送到 `git@github.com:Akashi86/Docx-AI-Refiner.git`。
- 每次对应用逻辑、提示词、处理流程、依赖、验证方式或已知问题做更新时，必须同步更新本文件，确保新线程可以无缝接手。

## 项目目标

这是一个 Streamlit 应用，用于处理用户上传的 Word 论文 `.docx`：

1. 读取 Word 文档正文段落。
2. 将正文逐段发送给 DeepSeek API，用指定提示词改写。
3. 将改写结果写回 Word。
4. 尽量不破坏原有 Word 论文格式。
5. 在网页上展示原文和改写后的对照清单，并支持下载 JSON/CSV 清单。
6. 目标不是“保证绕过检测器”，而是降低明显 AI 腔、模板句式、过度平滑和高风险段落模式。

根本需求是：**AI 改正文，但不要破坏 Word 论文原有格式，同时尽量降低 AIGC 检测风险。**

## 当前实现阶段

项目已经从“基础逐段润色”推进到：

**整段高成功率改写 + 段落风险扫描 + 高风险段落定向二次改写 + 改写对照清单 + 修订痕迹版下载。**

主要代码在 `app.py`。

## 当前关键实现

### 1. Word 读取与写回

- `.docx` 仍按 zip 读取，核心内容来自 `word/document.xml`。
- 扫描 `w:p` 段落。
- 读取段落内可写纯文本 run。
- 当前主流程是**整段文本发送给 AI**，而不是旧的多 `<seg id>` 精确映射。
- 写回时会移除该段的纯文本 runs，并用新文本生成新的 run。
- 会保留段落结构和尽量复用首个 run 的 `rPr` 样式。
- 会尝试保留原段落中加粗、斜体、下划线的术语：`formatted_terms_from_paragraph()` 会记录原格式术语，`split_text_by_terms()` 写回时尽量恢复对应 run 格式。
- 对复杂 inline 段落新增宽松写回：如果段落包含复杂 run、域、超链接等结构但仍有 `w:t` 文本节点，任务会使用 `write_mode = text_nodes_fallback`，把改写全文写入第一个可写文本节点，并清空其余文本节点；这样可避免高风险正文段落因复杂 inline 被整体跳过。

注意：当前实现已经不是严格的“只替换原有 `w:t.text`”。段落和大体样式保留，但段内局部 run 粒度可能变化。

### 2. 正文筛选

- 跳过常见非正文样式：`title`、`heading`、`toc`、`caption`、`header`、`footer`、`bibliography`、`reference`。
- 跳过短标题形态，如 `Abstract`、`Chapter`、`References`、`Appendix` 等。
- 按词数、标点、最小字符数过滤短段落。
- 旧逻辑会跳过复杂 inline 内容段落；最新逻辑不再直接跳过，而是对有文本节点的复杂段落使用 `text_nodes_fallback` 宽松写回。

### 3. 章节识别

- `extract_headings_from_docx()` 识别 Heading 1/2。
- UI 支持选择处理起止章节。
- `collect_tasks()` 在扫描段落时跟踪当前 Heading，并用 `detect_section_type()` 判断段落所属类型：
  - `abstract`
  - `introduction`
  - `literature_review`
  - `methodology`
  - `results_discussion`
  - `conclusion`
  - `general`

### 4. 提示词与强度

当前英文降 AI 逻辑不再强调“polish”或普通学术润色，而是强调：

- 像真实本科论文作者的二稿。
- 不要过度平滑、过度工整、处处对称。
- 不要仅做同义词替换。
- 保留事实、引用、术语、数据和信息顺序。
- 允许适度自然的不均匀节奏。

UI 里有“改写强度”：

- `标准降 AI`：基础温度 `0.55`
- `深度降 AI`：基础温度 `0.68`
- `最大强改写`：基础温度 `0.78`

### 5. 段落风险扫描

新增了通用风险扫描，不针对单篇论文写死。核心函数：

- `analyze_ai_risk(text, section_type)`
- `build_risk_instruction(risk_profile)`
- `task_temperature(base_temperature, risk_profile, section_type)`
- `should_force_risk_rewrite(...)`
- `build_force_rewrite_instruction(...)`

风险标签包括：

- `generic_study_opening`
  - 如 `This study`、`The findings`、`Existing research`、`To build upon` 等。
- `template_connectors`
  - 如 `moreover`、`furthermore`、`by contrast`、`on the other hand`、`according to` 等。
- `stock_interpretive_verbs`
  - 如 `shows`、`reveals`、`demonstrates`、`suggests`、`presents`、`offers` 等。
- `future_research_template`
  - 如 `future studies should`、`future research`、`longitudinal analysis would` 等。
- `results_data_template`
  - 表格、百分比、结果解释模板，如 `Tables 4.2 and 4.3 present`、`percentage-point gap`、`quantitative distribution`。
- `contribution_template`
  - 如 `offers several contributions`、`replicable analytical pathway`、`theoretical and practical implications`。
- `over_casual_naturalization`
  - 如 `real headaches`、`make or break`、`isn't a coincidence`、`zeroes in` 等，避免“AI 自然化口语腔”。
- `first_person_method`
  - 方法章不必要的一人称叙述，如 `I gathered`、`we used`。
- `abstract_noun_density`
  - 抽象名词堆叠过多，如 `strategy`、`framework`、`approach`、`distribution`、`analysis` 等。
- `formulaic_data_explanation`
  - 数字/百分比后接固定解释动词，如 `54.8% shows...`、`data suggests...`。

### 6. 高风险段落二次改写

当前逻辑会动态提高高风险段落温度，并在必要时二次调用 DeepSeek。

高优先级风险标签：

- `formulaic_data_explanation`
- `results_data_template`
- `future_research_template`
- `contribution_template`
- `generic_study_opening`
- `abstract_noun_density`

触发二次改写的典型条件：

- 改写后仍残留高优先级风险标签。
- 结果章或结论章仍残留真实风险标签。
- 原文和改写后相似度过高，且原段落有高优先级风险。
- 原段落风险很高，改写后仍未降到较低风险。

日志中出现：

```text
Paragraph X: detector-prone patterns remained, so a second rewrite was applied.
```

表示不是报错，而是该段第一次改写后仍残留检测器敏感模式，系统自动进行了第二轮定向改写。

### 7. 对照清单

`rewrite_report` 目前包含：

- `paragraph_index`
- `page`
- `original_text`
- `new_text`
- `old_diff_html`
- `new_diff_html`
- `status`
- `error`
- `section`
- `section_type`
- `write_mode`
- `risk_tags`
- `final_risk_tags`

JSON/CSV 下载也包含风险标签字段，便于分析哪类段落仍然残留风险。

### 8. 修订痕迹版

- `generate_tracked_changes_docx()` 使用 `docx_editor` 对原文和润色版生成修订痕迹文档。
- 如果修订痕迹生成失败，会保留普通润色版下载，不阻塞主流程。

### 9. 进度显示

- 进度条已经从单纯的“完成段落数 / 待处理段落数”改为阶段加权进度。
- 读取和扫描文档约占前 `6%`。
- AI 逐段改写约占 `6% -> 90%`，并显示“已完成 X/Y 段”。
- 打包润色版、生成修订痕迹版约占最后 `10%`。
- 这样可以减少高风险段落二次改写、文档打包、修订痕迹生成导致的“进度条不准”或“100% 后还在处理”的错觉。

## 针对样本文档的观察结论

样本文档位于 `word/`，该目录已在 `.gitignore` 中忽略。

曾经分析过：

- `高楠-论文.docx`
- `高楠-论文降低ai率版.docx`
- `润色版_高楠-论文.docx`
- `初版报告.pdf`
- `高楠-论文_AIGC统计报告.html`

检测结果对比：

- 初版报告整体 AIGC：`86.32%`
- 旧降 AI 版报告整体 AIGC：`65.71%`
- 新版文档尚不建议立刻付费整篇检测，原因是本地规则分析显示虽有继续改善，但仍有少数高风险段落残留。

旧版到新版的本地风险规则对照：

- 115 个正文长段。
- 平均风险标签：`1.04 -> 0.58`
- 抽象名词密度：`3.02 -> 2.52`
- 75 个段落明显改写。
- 42 个段落风险标签减少。
- 70 个段落持平。
- 3 个段落风险标签变多。

仍需重点处理的段落类型：

- 结果章里的表格/百分比解释。
- 结论章里的贡献模板。
- `future studies should` / `To build upon` 类未来研究模板。
- `This study offers...` 类贡献句。
- 改写后仍和原文高度相似的高风险段落。

用户在云端生成的最新版样本 `润色版_高楠-论文3.docx` 和 `rewrite_report.json` 显示：

- 共处理 `95` 段，`95` 段均为 `changed`。
- `76` 段使用 `rebuild_runs`。
- `19` 段使用 `text_nodes_fallback`。
- 之前漏掉的重点段落已经进入报告并被改写，包括：
  - `Tables 4.2 and 4.3...`
  - `First, regarding the quantitative distribution...`
  - `12.1-percentage-point gap...`
  - `This study offers several contributions...`
  - `To build upon the current findings...`
  - `Theoretically, this study seeks...`
- 但本地规则仍显示结果章和结论章有部分残留风险，常见于数据解释句、`This study...` 开头、表格说明句和抽象名词密度较高的段落。下一步如果继续推进，优先加强“结果章表格/百分比解释”和“结论章贡献/未来研究模板”的改写策略。

用户随后付费检测该最新版，整体疑似 AIGC 为 `77.1%`，说明前一轮策略存在方向性问题。检测报告显示：

- Abstract：`99.0%`
- 摘要：`84.0%`
- 绪论前言：`100.0%`
- Chapter One Introduction：`89.0%`
- Chapter Two Literature Review：`86.0%`
- Chapter Three Research Methodology：`71.0%`
- Chapter Four Results and Discussion：`61.0%`
- Chapter Five Conclusion：`78.0%`

复盘发现两个关键问题：

- 结构保护不足：`关键词` 被改成正文说明句，英文题名被改成正文段落，封面题名、原创性声明、致谢等前置文本也曾进入任务队列。
- 提示词过度强调“自然”和“二稿感”，导致 `kicked off`、`packed with`、`zeroes in`、`Let’s cut...`、反问句、缩写等口语/编辑腔表达，形成新的 AI 改写器风格。

已修正：

- 新增关键词、题名、中文长题名、原创性声明、致谢、前置信息保护规则。
- 写回前新增安全校验：如果保护性文本被扩写、标题被改成正文、AI 增加 Markdown 星号、或出现明显口语化改写器词组，则拒绝写回。
- 英文提示词从“自然化/像二稿”收回到“克制的本科论文风格”，明确禁止口语化、反问、缩写、杂志化表达、Markdown 和新增事实。

## 已完成事项

- 项目初始化为 Git 仓库并推送。
- 基础 Streamlit UI。
- DeepSeek API 调用。
- 有限重试，避免无限递归。
- 日志 HTML escape，避免注入风险。
- `.gitignore` 忽略：
  - `__pycache__/`
  - `*.py[cod]`
  - `.streamlit/secrets.toml`
  - `.env`
  - `tmp/`
  - `word/`
  - `streamlit*.log`
- 整段高成功率改写。
- 相邻同样式纯文本 run 合并。
- 章节范围选择。
- 对照清单展示和 JSON/CSV 下载。
- 改写状态区分：
  - `changed`
  - `unchanged`
  - `failed`
- 修订痕迹版 Word 下载。
- 英文降 AI 提示词强化。
- 改写强度 UI。
- 段落风险扫描与自动二次修复开关。
- 高风险段落动态升温。
- 高风险段落强制二次改写。
- 对照清单导出风险标签。
- 复杂 inline 段落宽松写回，避免结果章/结论章高风险段落因格式结构被跳过。
- 阶段加权进度条和进度说明，避免处理后段阶段显示不准。

## 踩坑点与处理经验

### 1. 多 seg 机制安全但成功率低

旧方案要求 AI 保留每个 `w:t` 的 `<seg id>`，格式保护强，但 Word 经常把普通段落拆成多个碎文本节点，AI 少返回一个片段就会失败。

经验：

- 这种方式适合极高格式保护场景。
- 普通论文正文更适合整段改写，提高成功率。

### 2. 整段写回更实用，但会牺牲 run 粒度

当前写回会重建纯文本 run，不是只改 `w:t.text`。

经验：

- 正文大多统一格式时，这个折中可接受。
- 复杂字段、链接、批注、脚注、编号等段落过去会被跳过；现在对含文本节点的复杂段落使用 `text_nodes_fallback`，保留原 XML 结构但会牺牲段内局部样式粒度。
- `rewrite_report.csv` 里的 `write_mode` 可用于确认某段是 `rebuild_runs` 还是 `text_nodes_fallback`。

### 3. “自然化”不等于降低检测风险

样本文档里发现，普通自然化可能把文本改成另一种 AI 腔：

- `This means that...`
- `When we look at...`
- `on the other hand`
- `That consistency isn't a coincidence`
- `real headaches`
- `make or break`
- `kicked off`
- `packed with`
- `zeroes in`
- `Let's cut...`

经验：

- 提示词不能只说“make it natural”。
- 要针对具体风险类型修复。
- 过度口语化也可能被检测器认为是 AI 改写器风格。
- 当前目标应是“克制、朴素、不新增内容的本科论文表达”，不是“更像口语的人类表达”。

### 4. 结构保护优先级高于改写覆盖率

最新版检测暴露出：如果应用误改关键词、题名、目录/前置元数据，检测结果会被明显拉高，而且文档内容也会被污染。

经验：

- `Keywords:` / `关键词：` / `關鍵詞：` 这类行必须跳过。
- 论文题名、封面中文长题名、原创性声明、致谢、署名和学校信息必须跳过。
- 没有句末标点、明显像题名的长文本，默认跳过比强行改写更安全。
- 即使某段误入任务队列，写回前也要用 `is_suspicious_expansion()` 拒绝“短标签/题名被扩写成正文”的返回。

### 5. 数据解释段和结论段是重灾区

容易高风险的句式：

- `Tables 4.2 and 4.3 present...`
- `First, regarding the quantitative distribution...`
- `The 12.1-percentage-point gap... shows...`
- `This study offers several contributions...`
- `To build upon the current findings, future studies should...`

经验：

- 结果章和结论章不能只靠通用提示词。
- 需要章节类型识别和专门风险规则。

### 6. 检测报告很有价值，但不要过早付费整篇检测

本地规则分析可以先判断是否仍有明显残留。

经验：

- 若最新版仍有明显高危模式，先继续修应用逻辑。
- 若检测平台支持章节检测，优先测 Abstract、Results/Discussion、Conclusion。
- 只能整篇检测时，等高风险规则稳定后再测。

### 7. PowerShell 中文显示可能乱码

曾出现 `AGENTS.md` 内容乱码、PowerShell 输出乱码等问题。

经验：

- 文件应保持 UTF-8。
- 不要只根据 PowerShell 乱码判断文件损坏。
- 但如果 `Get-Content -Raw` 读出的文件本身就是乱码，应重写为干净 UTF-8。

### 8. 本地 Streamlit 启动可能留下日志

曾因本地测试生成 `streamlit.err.log`、`streamlit.out.log`，文件被进程占用。

经验：

- 用户当前云端部署，通常不需要本地启动。
- `streamlit*.log` 已加入 `.gitignore`。
- 不要把本地运行日志提交。

### 9. 进度条不能只看段落完成数

高风险段落可能触发二次改写，主流程结束后还要打包 `.docx` 和生成修订痕迹版。

经验：

- 只用 `completed_tasks / len(tasks)` 会让用户感觉进度条不准。
- 新增耗时步骤时，要同步调整阶段权重和进度文案。
- 进度文案应明确当前处于扫描、AI 改写、打包还是修订痕迹生成阶段。

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

提交推送：

```powershell
git add app.py AGENTS.md .gitignore requirements.txt
git commit -m "中文提交信息"
git push
```

## 最近提交记录

- `d547fa8 增强英文降 AI 改写策略`
- `bf02601 新增段落风险扫描与定向改写`
- `70082b4 强化高风险段落二次改写`

## 待办事项

### 高优先级

- 用最新云端版本重新生成样本文档，然后用本地规则先复核高风险残留。
- 如果残留明显减少，再考虑付费整篇检测。
- 若检测后仍偏高，优先分析检测报告中的疑似片段汇总，而不是盲目改 prompt。
- 重点验证复杂 inline 宽松写回后的结果章/结论章段落是否真正进入 `rewrite_report`，尤其是 `Tables 4.2 and 4.3 present...`、`First, regarding...`、`The 12.1-percentage-point gap...`、`This study offers...`、`To build upon...`。
- 增加“检测报告返修模式”：
  1. 用户上传 AIGC 报告 HTML/PDF。
  2. 程序提取疑似片段。
  3. 在 Word 段落中匹配对应文本。
  4. 只对命中段落进行强改写。

### 中优先级

- 给日志信息全面改成中文，尤其是二次改写日志。
- UI 中展示每段风险标签和二次改写次数。
- 对 `risk_tags` 和 `final_risk_tags` 做更友好的中文解释。
- 给“高风险强制二次改写”增加统计摘要：多少段触发、触发原因是什么。
- 优化中文/中英混排差异高亮，目前主要是空格分词。

### 低优先级

- 增加“快速/平衡/强力”模式。目前逻辑偏强力，速度变慢但 token 消耗可接受，用户暂时不要求优化。
- 支持复制原文/复制改写后按钮。
- 支持只下载 changed 段落清单。
- 支持对 failed/unchanged 段落单独重试。

## 注意事项

- 不要承诺绕过或保证通过 AIGC 检测；检测器不稳定，且不同平台规则不同。
- 应描述为“降低明显 AI 腔和检测器敏感模式”。
- 修改 `.docx` 时优先保留 XML 结构，宁愿跳过复杂段落，也不要破坏文档。
- 如果工作涉及应用逻辑更新，完成后必须同步更新本文件。
