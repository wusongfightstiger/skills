# PPT 中译英 Skill 设计文档

> 日期: 2026-03-29
> 状态: 已批准

## 1. 概述

一个 Claude Code skill，将中文 PowerPoint 翻译为英文，保持原始格式。面向学术/技术场景（课件、论文答辩、技术报告等）。

**调用方式：**
```
/ppt-translator-opus <pptx路径> [--engine claude|minimax|claude-api] [--glossary <术语表路径>] [--output <输出路径>]
```

## 2. 整体架构

```
用户: /ppt-translator-opus 第1章.pptx --engine minimax
                │
        ┌───────┴───────┐
        │   SKILL.md     │  ← Claude Code 的行为指令
        │   编排整个流程   │
        └───────┬───────┘
                │
        ┌───────┴───────┐
        │  Python 包      │  ppt_translate/
        │  (CLI 入口)     │
        │                 │
        │  两种运行模式:   │
        │  ├─ api 模式    │ → 完整自动运行，内部调 API
        │  └─ claude 模式 │ → 提取 → Claude 翻译 → 回填
        └───────────────┘
```

### 运行模式

- **claude 模式（默认）**：Python 负责 PPTX 拆解/重组，翻译由 Claude Code 在对话中完成。流程：
  1. Python `extract` → `slides_text.json`
  2. Claude 读取 JSON，逐页翻译，写入 `slides_translated.json`
  3. Python `apply` → 回填 PPTX
- **api 模式**（minimax/claude-api）：Python 全自动完成，Claude Code 只启动和报告结果

## 3. PPTX 处理引擎

以 `python-pptx` 为主，SmartArt 等不支持的元素降级到 XML 操作。

### 支持的元素类型

| 元素 | 库/方法 | 说明 |
|------|---------|------|
| 文本框 (`p:sp`) | `python-pptx` Shape API | 最常见，直接用高层 API |
| 表格 (`p:graphicFrame`) | `python-pptx` Table API | 逐单元格提取/回填 |
| 分组形状 (`p:grpSp`) | `python-pptx` GroupShapes + 递归 | 递归遍历组内所有子形状 |
| 备注 (`p:notes`) | `python-pptx` NotesSlide API | 每张幻灯片的备注页 |
| SmartArt | 降级到 XML 操作 | 直接解析 `dgm:` 命名空间下的文本节点 |

### 文本提取数据结构

```json
{
  "slide_number": 1,
  "elements": [
    {
      "id": "s1_shape_3",
      "type": "textbox",
      "runs": [
        {"index": 0, "text": "电路", "bold": true, "font_size": 24},
        {"index": 1, "text": "的基本原理", "bold": false, "font_size": 24}
      ]
    },
    {
      "id": "s1_table_1_r0c0",
      "type": "table_cell",
      "runs": [
        {"index": 0, "text": "参数名称", "bold": true, "font_size": 14}
      ]
    },
    {
      "id": "s1_note",
      "type": "note",
      "runs": [
        {"index": 0, "text": "本页讲解基本概念", "bold": false, "font_size": 12}
      ]
    }
  ]
}
```

### 智能格式映射

翻译时保留 run 结构信息，LLM 返回相同 JSON 结构，只替换 `text` 字段。按 `id` 和 `index` 匹配回填，格式（粗体、字号、颜色等）保持不变。

### 回填逻辑

- 按 `id` 匹配原始元素，按 `index` 匹配 run
- 每个 run 的文本替换，格式保持不变
- 仅将东亚字体（`ea`）改为 Arial/Calibri，`latin` 和 `cs` 字体保持原样

### AutoFit 处理

- 计算译文与原文的字符数比值，按比例缩放字号
- 最小字号下限 8pt，超出则启用 `python-pptx` 的 autofit 缩放

## 4. 翻译引擎（可插拔）

### 引擎接口

```python
class TranslationEngine(ABC):
    @abstractmethod
    def translate_slide(self, slide_data: dict, glossary: list[dict]) -> dict:
        """翻译一张幻灯片的全部元素，返回译文 JSON"""
        pass
```

### 三个实现

**ClaudeCodeEngine：**
- 不调任何 API
- 将 JSON 写入临时文件，返回特殊标记让 SKILL.md 编排 Claude 翻译
- 整个 PPT 分页处理，每页一轮交互

**MiniMaxEngine：**
- 调用 `https://api.minimaxi.com/anthropic/v1/messages`
- 模型可配置（默认 `MiniMax-M2.7`）
- 需要 `MINIMAX_API_KEY` 环境变量

**ClaudeAPIEngine：**
- 调用 Anthropic Messages API
- 模型可配置（默认 `claude-sonnet-4-6-20250514`）
- 需要 `ANTHROPIC_API_KEY` 环境变量

### 并行调度（API 模式）

```python
async def translate_all_slides(slides, engine, max_concurrent=5):
    semaphore = asyncio.Semaphore(max_concurrent)

    async def translate_one(slide):
        async with semaphore:
            return await engine.translate_slide(slide, glossary)

    tasks = [translate_one(s) for s in slides]
    results = await asyncio.gather(*tasks, return_exceptions=True)
```

- 默认 5 并发
- 收到 429 → 读取 `Retry-After`，指数退避（1s, 2s, 4s），最多重试 3 次
- 动态降级：连续 3 次 429 → 并发数减半
- 单页失败不阻塞，最终汇报失败列表

### Claude Code 模式流程编排

由 SKILL.md 指导 Claude Code 执行：

1. 运行 `ppt-translator-opus extract <path>` → `slides_text.json`
2. 读取 `slides_text.json`
3. 如有术语表，读取术语表
4. 逐页翻译：
   - 读取该页 JSON
   - 匹配命中的术语
   - 构造 prompt，Claude 自己翻译
   - 将译文追加写入 `slides_translated.json`
5. 运行 `ppt-translator-opus apply <path> <translated_json> --output <output>`
6. 报告结果

## 5. 术语表

### 格式

简化的 CSV，面向通用学术场景：

```csv
中文,英文,领域
电路,circuit,电气工程
赫罗图,Hertzsprung-Russell diagram,天文学
```

- `领域` 可选，用于未来按领域筛选
- 通过 `--glossary` 参数指定路径，不指定则不使用术语表

### 术语注入方式

不做预替换。从当前幻灯片文本中匹配命中的术语，只把命中的术语作为上下文传给 LLM：

```python
def extract_relevant_terms(slide_text: str, glossary: list) -> list:
    return [term for term in glossary if term['中文'] in slide_text]
```

## 6. Prompt 工程

### 翻译 Prompt 模板

```
你是一位专业的学术/技术文档翻译专家。将以下幻灯片内容从中文翻译为英文。

要求：
1. 保持 JSON 结构不变，只翻译 "text" 字段
2. 保持每个元素的 runs 数量不变
3. 学术/技术术语翻译准确，语言正式简洁
4. 数字、公式、变量名保持原样

术语参考（请遵循）：
{matched_terms}

幻灯片内容：
{slide_json}

只返回翻译后的 JSON，不要解释。
```

### Prompt 分层

- **系统 prompt**：角色定义 + 通用规则（不变）
- **术语段**：按页动态生成（只含命中术语）
- **内容段**：当前幻灯片的 JSON

## 7. 错误处理

### 三级容错

| 级别 | 场景 | 处理 |
|------|------|------|
| Run 级 | 某个 run 的译文回填失败 | 该 run 保留原文，其余 run 正常回填 |
| 元素级 | LLM 返回的某个元素 id 匹配不上 | 跳过该元素，保留原文 |
| 幻灯片级 | 整页翻译失败（API 超时/重试耗尽/JSON 解析失败） | 整页保留原文，记录到失败列表 |

核心原则：**永远不会产出一个损坏的 PPTX**。任何失败都回退到原文。

### JSON 校验

```python
def validate_translation(original, translated):
    - 元素数量一致
    - 每个元素的 id 匹配
    - 每个元素的 runs 数量一致
    - text 字段非空
    # 校验失败 → 重试一次 → 仍失败则保留原文
```

## 8. 输出

- 默认输出路径：`<原文件名>_en.pptx`（同目录）
- 可通过 `--output` 指定

### 翻译完成摘要报告

```
翻译完成！

输入:  第1章.pptx (25 张幻灯片)
输出:  第1章_en.pptx
引擎:  minimax (MiniMax-M2.7)
耗时:  2 分 34 秒

结果:
  ✓ 成功: 23 张幻灯片 (142 个元素)
  ✗ 失败: 2 张幻灯片 (第 8, 15 页)
  - 第 8 页: API 超时
  - 第 15 页: JSON 解析失败

术语表命中: 47 个术语
```

## 9. 文件结构

```
~/.claude/skills/ppt-translator-opus/
├── SKILL.md                     # Skill 定义文件
├── pyproject.toml               # Python 包配置
├── src/ppt_translator_opus/
│   ├── __init__.py
│   ├── cli.py                   # Click CLI 入口 (extract / apply / translate)
│   ├── pptx_engine.py           # PPTX 解析/回填（python-pptx + XML 降级）
│   ├── engines/
│   │   ├── __init__.py
│   │   ├── base.py              # TranslationEngine ABC
│   │   ├── minimax.py           # MiniMax 实现
│   │   └── claude_api.py        # Anthropic API 实现
│   ├── glossary.py              # 术语表加载与匹配
│   ├── prompt.py                # Prompt 模板构建
│   └── utils.py                 # 并行调度、重试、校验
└── tests/
    ├── test_pptx_engine.py
    ├── test_engines.py
    ├── test_glossary.py
    └── test_prompt.py
```

## 10. CLI 子命令

| 命令 | 用途 | 模式 |
|------|------|------|
| `ppt-translator-opus translate` | 全自动翻译（API 模式专用） | minimax / claude-api |
| `ppt-translator-opus extract` | 只提取文本 JSON | claude 模式第一步 |
| `ppt-translator-opus apply` | 读取译文 JSON，回填 PPTX | claude 模式最后一步 |

## 11. 依赖

- `python-pptx` — PPTX 读写
- `click` — CLI 框架
- `httpx` — 异步 HTTP（API 调用）
- `lxml` — XML 处理（SmartArt 降级）
