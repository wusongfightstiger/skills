---
name: ppt-translator-opus
description: "将中文 PPT 翻译为英文 PPT，保持原始格式。支持文本框、表格、分组形状、SmartArt、备注。支持 Claude Code 自身翻译和外部 API（MiniMax、Claude API）。"
---

# PPT 中译英翻译器

将中文 PowerPoint 翻译为英文，保持原始格式和排版。

## 用法

```
/ppt-translator-opus <pptx路径> [--engine claude|minimax|claude-api] [--glossary <术语表路径>] [--no-glossary] [--output <输出路径>]
```

**参数说明：**
- `<pptx路径>`: 必需，待翻译的 PPTX 文件路径
- `--engine`: 翻译引擎，默认 `claude`
  - `claude`: Claude Code 自身翻译（无需额外 API）
  - `minimax`: MiniMax API（需要 MINIMAX_API_KEY 环境变量）
  - `claude-api`: Anthropic API（需要 ANTHROPIC_API_KEY 环境变量）
- `--glossary`: 可选，指定自定义术语表路径（CSV 或 TXT），覆盖默认术语表
- `--no-glossary`: 禁用术语表
- `--output`: 可选，输出路径（默认 `<原文件名>_en.pptx`）

**术语表：**
- 默认自动加载 `~/.claude/skills/ppt-translator-opus/glossary.csv`
- 使用 `--glossary <path>` 指定其他术语表
- 使用 `--no-glossary` 完全禁用术语表
- 术语表可手动编辑，CSV 格式：`中文,英文,领域`
- 详细规则见 `docs/glossary-guide.md`

## 执行流程

首先确保依赖已安装：
```bash
pip3 install -e ~/.claude/skills/ppt-translator-opus
```

### API 模式 (minimax / claude-api)

API 模式下，Python 脚本全自动完成翻译，Claude Code 只需启动命令并报告结果。

```bash
ppt-translator-opus translate <pptx路径> --engine <minimax|claude-api> [--glossary <path>] [--output <path>]
```

直接运行上面的命令即可。脚本会自动完成提取、翻译、回填，并输出摘要报告。

### Claude 模式 (默认)

Claude 模式下，Python 负责 PPTX 拆解/重组，翻译由你（Claude）在对话中完成。按以下步骤执行：

**第一步：提取文本**
```bash
ppt-translator-opus extract <pptx路径> --output /tmp/ppt_translate_work/slides_text.json
```

**第二步：读取提取的 JSON**

用 Read 工具读取 `/tmp/ppt_translate_work/slides_text.json`。

**第三步：读取术语表**

除非用户指定了 `--no-glossary`，否则读取术语表：
- 用户指定了 `--glossary <path>` → 读取该路径
- 未指定 → 读取默认术语表 `~/.claude/skills/ppt-translator-opus/glossary.csv`

**第四步：逐页翻译**

对 JSON 中的每一张幻灯片（每个 slide 对象），执行翻译：

1. 取出当前幻灯片的 JSON 数据
2. 如有术语表，找出当前页中出现的术语
3. 翻译该页所有 elements 中 runs 的 text 字段（中文 → 英文）
4. **保持 JSON 结构完全不变**：不改变 id、type、runs 数量、index，只替换 text 值
5. 翻译要求：
   - 学术/技术术语准确
   - 语言正式简洁
   - 数字、公式、变量名保持原样
   - 如果原文已是英文，保持不变
   - 如有术语表命中，遵循术语表翻译

翻译完成后，将所有翻译结果组成一个 JSON 数组，写入 `/tmp/ppt_translate_work/slides_translated.json`。

格式示例：
```json
[
  {
    "slide_number": 1,
    "elements": [
      {
        "id": "s1_shape_1",
        "type": "textbox",
        "runs": [
          {"index": 0, "text": "Circuit Fundamentals"}
        ]
      }
    ]
  }
]
```

**第五步：回填 PPTX**
```bash
ppt-translator-opus apply <pptx路径> /tmp/ppt_translate_work/slides_translated.json --output <输出路径>
```

**第六步：报告结果**

告诉用户翻译完成，输出文件路径，以及成功/失败统计。

## 注意事项

- 翻译时严格保持 JSON 结构不变，只改 text 字段
- 每张幻灯片独立翻译，不要跨页合并
- 如果某页翻译失败，保留原文继续处理下一页
- 大型 PPT（超过 30 页）时，可以分批处理（每 10 页一批）以避免上下文过长

## 进度监控

**API 模式：** 内置自动进度报告，每 10 分钟输出一次进度摘要（成功/失败/剩余数量、当前时间、预计剩余时间）。

**Claude 模式：** 在逐页翻译过程中，每翻译完 10 页，主动向用户报告一次进度，格式：

```
进度报告 [HH:MM]
  ✓ 成功: N 页
  ✗ 失败: N 页
  … 剩余: N 页
  预计还需: X 分钟
```
