"""Translation prompt templates."""

import json


SYSTEM_PROMPT = """你是一位专业的学术/技术文档翻译专家。你的任务是将中文幻灯片内容翻译为英文。

核心要求：
1. 保持 JSON 结构不变，只翻译 "text" 字段
2. 保持每个元素的 runs 数量不变，不要增加或减少 runs
3. 学术/技术术语翻译准确，语言正式简洁
4. 数字、公式、变量名、单位保持原样
5. 如果原文已经是英文，保持不变
6. 只返回翻译后的 JSON，不要任何解释或 markdown 包裹"""


def build_slide_prompt(slide_data: dict, matched_terms: list[dict] | None = None) -> str:
    """Build a translation prompt for one slide.

    Args:
        slide_data: The slide dict with slide_number and elements.
        matched_terms: List of matched glossary terms for this slide.

    Returns:
        The user prompt string.
    """
    parts = []

    parts.append(f"翻译以下第 {slide_data['slide_number']} 张幻灯片的内容（中文 → 英文）。")

    if matched_terms:
        parts.append("\n术语参考（请遵循）：")
        for term in matched_terms:
            parts.append(f"- {term['zh']} → {term['en']}")

    # Build a simplified version of slide data for translation (only id, type, runs with index+text)
    simplified = {
        "slide_number": slide_data["slide_number"],
        "elements": [],
    }
    for elem in slide_data["elements"]:
        simplified_elem = {
            "id": elem["id"],
            "type": elem["type"],
            "runs": [{"index": r["index"], "text": r["text"]} for r in elem["runs"]],
        }
        simplified["elements"].append(simplified_elem)

    parts.append(f"\n幻灯片内容：\n{json.dumps(simplified, ensure_ascii=False, indent=2)}")

    return "\n".join(parts)
