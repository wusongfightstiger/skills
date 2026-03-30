"""Glossary loader and term matching."""

import csv
from pathlib import Path


def load_glossary(glossary_path: str) -> list[dict]:
    """Load a glossary from CSV or TXT file.

    CSV format: 中文,英文,领域
    TXT format: 中文<tab>英文 (one pair per line)

    Returns list of {"zh": ..., "en": ..., "domain": ...} dicts.
    """
    path = Path(glossary_path)
    terms = []

    if path.suffix.lower() == ".csv":
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Support both column naming conventions:
                #   新格式: 中文,英文,领域
                #   旧格式: 中文术语,英文翻译,领域,添加日期,是否已确认
                zh = (row.get("中文") or row.get("中文术语") or "").strip()
                en = (row.get("英文") or row.get("英文翻译") or "").strip()
                domain = (row.get("领域") or "").strip()
                if zh and en:
                    terms.append({"zh": zh, "en": en, "domain": domain})
    else:
        # TXT or other: assume tab-separated or space-separated
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.split("\t") if "\t" in line else line.split(None, 1)
                if len(parts) >= 2:
                    zh = parts[0].strip()
                    en = parts[1].strip()
                    if zh and en:
                        terms.append({"zh": zh, "en": en, "domain": ""})

    return terms


def extract_relevant_terms(slide_text: str, glossary: list[dict]) -> list[dict]:
    """Return only terms that appear in the given slide text."""
    return [term for term in glossary if term["zh"] in slide_text]
