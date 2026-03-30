# PPT Translator Opus — 实现计划

> 日期: 2026-03-29
> 设计文档: `docs/superpowers/specs/2026-03-29-ppt-translate-design.md`
> Skill 路径: `~/.claude/skills/ppt-translator-opus/`

## 实现步骤

### Step 1: 项目脚手架

创建目录结构和基础配置文件。

**文件：**
- `~/.claude/skills/ppt-translator-opus/pyproject.toml`
- `~/.claude/skills/ppt-translator-opus/src/ppt_translator_opus/__init__.py`

**pyproject.toml 要点：**
```toml
[project]
name = "ppt-translator-opus"
version = "0.1.0"
dependencies = ["python-pptx", "click", "httpx", "lxml"]

[project.scripts]
ppt-translator-opus = "ppt_translator_opus.cli:main"
```

**完成标准：** `pip install -e .` 成功，`ppt-translator-opus --help` 可运行。

---

### Step 2: PPTX 文本提取引擎

实现从 PPTX 中提取所有文本元素（含 run 结构和格式信息）。

**文件：** `src/ppt_translator_opus/pptx_engine.py`

**实现要点：**
1. `extract_slides(pptx_path) -> list[dict]`
   - 用 `python-pptx` 打开文件
   - 遍历每张幻灯片
   - 提取文本框：遍历 `slide.shapes`，对 `has_text_frame` 的 shape 提取 paragraphs → runs
   - 提取表格：检测 `shape.has_table`，遍历 rows → cells → paragraphs → runs
   - 提取分组形状：递归遍历 `shape.shapes`（GroupShapes）
   - 提取备注：`slide.notes_slide.notes_text_frame` → paragraphs → runs
   - 提取 SmartArt：降级到 XML，在 shape._element 中查找 `a:t` 文本节点
2. 每个元素生成唯一 id：`s{slide_num}_{type}_{index}`
3. 每个 run 记录：index, text, bold, italic, font_size, font_color
4. 输出为设计文档中定义的 JSON 结构
5. 跳过纯空白/空文本的元素

**完成标准：** 能对 `PPT翻译/第1章.pptx` 提取出完整的 JSON，包含文本框、表格、备注。

---

### Step 3: PPTX 译文回填引擎

实现将翻译后的 JSON 写回 PPTX。

**文件：** `src/ppt_translator_opus/pptx_engine.py`（追加）

**实现要点：**
1. `apply_translations(pptx_path, translations, output_path)`
   - 用 `python-pptx` 打开原始 PPTX
   - 按 id 匹配元素，按 index 匹配 run
   - 替换 run.text，保持其余格式属性不变
   - 东亚字体 (ea) 改为 Arial
   - AutoFit：计算 len(translated)/len(original) 比值，按比例缩放字号，下限 8pt
   - SmartArt 回填：通过 XML 直接修改 `a:t` 节点
2. 三级容错：
   - run 级：单个 run 回填失败 → 保留原文
   - 元素级：id 匹配不上 → 跳过
   - 幻灯片级：整页异常 → 保留原始幻灯片
3. 保存到 output_path

**完成标准：** 手动构造一份译文 JSON，回填后生成的 PPTX 可正常打开，格式保持。

---

### Step 4: 术语表模块

**文件：** `src/ppt_translator_opus/glossary.py`

**实现要点：**
1. `load_glossary(csv_path) -> list[dict]` — 加载 CSV（中文, 英文, 领域）
2. `extract_relevant_terms(slide_text, glossary) -> list[dict]` — 匹配当前幻灯片中出现的术语
3. 支持 CSV 和 TXT（`中文\t英文` 格式）两种输入

**完成标准：** 能正确加载 `电路术语表.csv`，对一段文本返回命中术语列表。

---

### Step 5: Prompt 模板模块

**文件：** `src/ppt_translator_opus/prompt.py`

**实现要点：**
1. `build_system_prompt() -> str` — 固定的系统角色 prompt
2. `build_slide_prompt(slide_json, matched_terms) -> str` — 按页构建翻译 prompt
3. 术语段只包含当前页命中的术语
4. 输出要求：只返回 JSON，保持结构不变

**完成标准：** 生成的 prompt 结构正确，术语动态注入。

---

### Step 6: 翻译引擎 — 基类 + MiniMax

**文件：**
- `src/ppt_translator_opus/engines/base.py`
- `src/ppt_translator_opus/engines/minimax.py`
- `src/ppt_translator_opus/engines/__init__.py`

**实现要点：**
1. `TranslationEngine` ABC：`translate_slide(slide_data, glossary) -> dict`
2. `MiniMaxEngine`：
   - 调用 `https://api.minimaxi.com/anthropic/v1/messages`，模型 `MiniMax-M2.7`
   - 用 `httpx.AsyncClient` 异步调用
   - 解析 Anthropic 格式响应，提取 text content
   - 从响应中解析 JSON（处理 markdown code fence 包裹的情况）
3. JSON 校验：`validate_translation(original, translated)` — 检查元素数、id 匹配、runs 数量

**完成标准：** 设置 MINIMAX_API_KEY 后，能翻译一张幻灯片并返回合法 JSON。

---

### Step 7: 翻译引擎 — Claude API

**文件：** `src/ppt_translator_opus/engines/claude_api.py`

**实现要点：**
1. `ClaudeAPIEngine`：
   - 调用 Anthropic Messages API
   - 默认模型 `claude-sonnet-4-6-20250514`，可通过参数覆盖
   - 用 `httpx.AsyncClient`
   - JSON 解析与校验同 MiniMax

**完成标准：** 设置 ANTHROPIC_API_KEY 后，能翻译一张幻灯片。

---

### Step 8: 并行调度与重试

**文件：** `src/ppt_translator_opus/utils.py`

**实现要点：**
1. `translate_all_slides(slides, engine, glossary, max_concurrent=5) -> list[dict]`
   - asyncio.Semaphore 控制并发
   - asyncio.gather 并行执行
2. 重试逻辑：
   - 429 → 读取 Retry-After，指数退避（1s, 2s, 4s），最多 3 次
   - 连续 3 次 429 → 并发数减半
3. 异常处理：单页失败返回 None，不阻塞其他页
4. `validate_translation(original, translated)` — JSON 结构校验

**完成标准：** 能并行翻译多张幻灯片，429 时正确退避。

---

### Step 9: CLI 入口

**文件：** `src/ppt_translator_opus/cli.py`

**实现要点：**
1. Click group + 3 个子命令：
   - `extract` — 提取文本 JSON 到临时目录
   - `apply` — 读取译文 JSON 回填 PPTX
   - `translate` — 全自动翻译（API 模式）
2. `translate` 命令：
   - `--engine` 选择引擎（minimax/claude-api）
   - `--glossary` 可选术语表
   - `--output` 输出路径（默认 `<name>_en.pptx`）
   - 调用 extract → translate_all_slides → apply
   - 输出摘要报告（成功/失败数、耗时、术语命中数）
3. `extract` 命令：
   - 输出 JSON 到 stdout 或 `--output` 指定的文件
4. `apply` 命令：
   - 读取 JSON 文件，回填 PPTX

**完成标准：** `ppt-translator-opus translate 第1章.pptx --engine minimax` 能端到端运行。

---

### Step 10: SKILL.md

**文件：** `~/.claude/skills/ppt-translator-opus/SKILL.md`

**实现要点：**

```markdown
---
name: ppt-translator-opus
description: "将中文 PPT 翻译为英文 PPT，保持原始格式。支持文本框、表格、分组形状、SmartArt、备注。"
---
```

内容：
1. 用法说明（参数、示例）
2. API 模式流程：安装依赖 → 运行 translate 命令 → 报告结果
3. Claude 模式流程：安装依赖 → extract → 逐页翻译（Claude 自己做） → apply → 报告结果
4. Claude 模式下的翻译 prompt 模板（内嵌在 SKILL.md 中，供 Claude 参考）
5. 错误处理说明

**完成标准：** `/ppt-translator-opus 第1章.pptx` 在 Claude Code 中可正常触发。

---

### Step 11: 集成测试

用 `PPT翻译/第1章.pptx` 做端到端测试：

1. **Claude 模式测试：** `/ppt-translator-opus 第1章.pptx`
2. **MiniMax 模式测试：** `/ppt-translator-opus 第1章.pptx --engine minimax`
3. **带术语表测试：** `/ppt-translator-opus 第1章.pptx --engine minimax --glossary PPT翻译/电路术语表.csv`
4. 验证输出 PPTX：格式保持、文本已翻译、无损坏

**完成标准：** 三种模式均能正常产出翻译后的 PPTX。

---

## 实现顺序与依赖

```
Step 1 (脚手架)
  └→ Step 2 (提取) + Step 4 (术语表) + Step 5 (Prompt)  [可并行]
       └→ Step 3 (回填)
            └→ Step 6 (MiniMax 引擎) + Step 7 (Claude API 引擎)  [可并行]
                 └→ Step 8 (并行调度)
                      └→ Step 9 (CLI)
                           └→ Step 10 (SKILL.md)
                                └→ Step 11 (集成测试)
```

## 预计工作量

11 个步骤，核心代码约 6 个 Python 文件 + 1 个 SKILL.md。
