"""CLI entry point for ppt-translator-opus."""

import asyncio
import json
import sys
import time
from pathlib import Path

import click

from .pptx_engine import extract_slides, apply_translations
from .glossary import load_glossary
from .utils import translate_all_slides, merge_translations

# Default glossary: glossary.csv in the skill directory
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
DEFAULT_GLOSSARY = SKILL_DIR / "glossary.csv"


def _resolve_glossary(glossary: str | None, no_glossary: bool) -> str | None:
    """Resolve glossary path: --no-glossary disables, --glossary overrides, else use default."""
    if no_glossary:
        return None
    if glossary:
        return glossary
    if DEFAULT_GLOSSARY.exists():
        return str(DEFAULT_GLOSSARY)
    return None


@click.group()
def main():
    """PPT Translator Opus - Translate Chinese PPT to English."""
    pass


@main.command()
@click.argument("pptx_path", type=click.Path(exists=True))
@click.option("--output", "-o", type=click.Path(), help="Output JSON path (default: stdout)")
def extract(pptx_path, output):
    """Extract text from PPTX as JSON."""
    slides = extract_slides(pptx_path)

    # Filter out slides with no elements
    slides = [s for s in slides if s["elements"]]

    data = json.dumps(slides, ensure_ascii=False, indent=2)

    if output:
        Path(output).write_text(data, encoding="utf-8")
        total_elements = sum(len(s["elements"]) for s in slides)
        click.echo(
            f"Extracted {len(slides)} slides ({total_elements} elements) → {output}"
        )
    else:
        click.echo(data)


@main.command()
@click.argument("pptx_path", type=click.Path(exists=True))
@click.argument("translations_json", type=click.Path(exists=True))
@click.option("--output", "-o", type=click.Path(), help="Output PPTX path")
def apply(pptx_path, translations_json, output):
    """Apply translated JSON back to PPTX."""
    if not output:
        p = Path(pptx_path)
        output = str(p.parent / f"{p.stem}_en{p.suffix}")

    translations = json.loads(Path(translations_json).read_text(encoding="utf-8"))
    stats = apply_translations(pptx_path, translations, output)

    click.echo(f"Applied translations → {output}")
    click.echo(f"  Runs: {stats['runs_ok']} ok, {stats['runs_fail']} failed")
    click.echo(f"  Elements: {stats['elements_ok']} ok, {stats['elements_skip']} skipped")
    click.echo(f"  Slides: {stats['slides_ok']} ok, {stats['slides_fail']} failed")


@main.command()
@click.argument("pptx_path", type=click.Path(exists=True))
@click.option(
    "--engine", "-e",
    type=click.Choice(["minimax", "claude-api"]),
    default="minimax",
    help="Translation engine",
)
@click.option("--glossary", "-g", type=click.Path(exists=True), help="Glossary CSV/TXT path (default: 电路术语表.csv in skill dir)")
@click.option("--no-glossary", is_flag=True, default=False, help="Disable glossary")
@click.option("--output", "-o", type=click.Path(), help="Output PPTX path")
@click.option("--concurrency", "-c", type=int, default=5, help="Max concurrent API calls")
def translate(pptx_path, engine, glossary, no_glossary, output, concurrency):
    """Translate PPTX fully automatically using an API engine."""
    if not output:
        p = Path(pptx_path)
        output = str(p.parent / f"{p.stem}_en{p.suffix}")

    start_time = time.time()

    # 1. Extract
    click.echo(f"Extracting text from {pptx_path}...")
    slides = extract_slides(pptx_path)
    slides_with_content = [s for s in slides if s["elements"]]
    total_elements = sum(len(s["elements"]) for s in slides_with_content)
    click.echo(f"  Found {len(slides_with_content)} slides with content ({total_elements} elements)")

    # 2. Load glossary
    glossary_terms = []
    glossary_path = _resolve_glossary(glossary, no_glossary)
    if glossary_path:
        glossary_terms = load_glossary(glossary_path)
        click.echo(f"  Loaded {len(glossary_terms)} glossary terms from {Path(glossary_path).name}")

    # 3. Create engine
    if engine == "minimax":
        from .engines.minimax import MiniMaxEngine
        eng = MiniMaxEngine()
    else:
        from .engines.minimax import ClaudeAPIEngine
        eng = ClaudeAPIEngine()

    click.echo(f"Translating with {eng.name()}, concurrency={concurrency}...")

    # 4. Translate
    translated = asyncio.run(
        translate_all_slides(slides_with_content, eng, glossary_terms, concurrency)
    )

    # 5. Merge (fallback to original on failure)
    merged = merge_translations(slides_with_content, translated)

    # Rebuild full slides list (including empty ones)
    full_translations = []
    content_idx = 0
    for s in slides:
        if s["elements"]:
            full_translations.append(merged[content_idx])
            content_idx += 1
        else:
            full_translations.append(s)

    # 6. Apply
    click.echo(f"Applying translations to {output}...")
    stats = apply_translations(pptx_path, full_translations, output)

    # 7. Report
    elapsed = time.time() - start_time
    success_slides = sum(1 for t in translated if t is not None)
    failed_slides = sum(1 for t in translated if t is None)
    failed_indices = [
        slides_with_content[i]["slide_number"]
        for i, t in enumerate(translated)
        if t is None
    ]

    click.echo(f"\n翻译完成！")
    click.echo(f"")
    click.echo(f"输入:  {Path(pptx_path).name} ({len(slides)} 张幻灯片)")
    click.echo(f"输出:  {Path(output).name}")
    click.echo(f"引擎:  {eng.name()}")
    click.echo(f"耗时:  {elapsed:.0f} 秒")
    click.echo(f"")
    click.echo(f"结果:")
    click.echo(f"  ✓ 成功: {success_slides} 张幻灯片 ({stats['elements_ok']} 个元素)")
    if failed_slides > 0:
        click.echo(f"  ✗ 失败: {failed_slides} 张幻灯片 (第 {', '.join(map(str, failed_indices))} 页)")

    if glossary_terms:
        # Count how many terms were actually matched across all slides
        all_text = " ".join(
            r["text"]
            for s in slides_with_content
            for e in s["elements"]
            for r in e["runs"]
        )
        from .glossary import extract_relevant_terms
        matched = extract_relevant_terms(all_text, glossary_terms)
        click.echo(f"  术语表命中: {len(matched)} 个术语")


if __name__ == "__main__":
    main()
