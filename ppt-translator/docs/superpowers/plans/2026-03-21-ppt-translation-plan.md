# PPT翻译工具实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 创建中文PPT翻译成英文PPT的工具，支持术语表管理和MiniMax API翻译

**Architecture:**
- CLI工具 + Claude Skill结合
- unpack/edit/pack工作流处理PPTX
- 文本框级处理：每个文本框作为独立翻译单位
- 字体处理：翻译后文字设置为西文字体
- autoFit处理：处理字符扩张问题

**Tech Stack:** Python 3.10+, MiniMax API, zipfile/xml.etree, CSV, pypinyin

---

## 文件结构

```
ppt-translator/
├── src/
│   └── ppt_translator/
│       ├── __init__.py
│       ├── config.py          # 配置管理，API密钥
│       ├── terminology.py      # 术语表读写
│       ├── translator.py       # MiniMax API调用
│       ├── ppt_handler.py      # PPT unpack/edit/pack，字体处理
│       ├── text_box.py         # 文本框模型
│       └── cli.py              # CLI入口
├── tests/                      # 测试目录
├── 电路术语表.csv
├── docs/
│   └── superpowers/
│       ├── specs/
│       └── plans/
└── README.md
```

---

## Task 1: 项目初始化

**Files:**
- Create: `src/ppt_translator/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`

- [ ] **Step 1: Create src/ppt_translator/__init__.py**

```python
"""PPT Translator - Translate Chinese PPT to English."""

__version__ = "0.1.0"
```

- [ ] **Step 2: Create tests/__init__.py**

```python
"""Tests for PPT Translator."""
```

- [ ] **Step 3: Create tests/conftest.py**

```python
import pytest
from pathlib import Path
import csv
from datetime import date


@pytest.fixture
def test_glossary_csv(tmp_path):
    """Create a temporary test glossary CSV."""
    csv_path = tmp_path / "test_glossary.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["中文术语", "英文翻译", "领域", "添加日期", "是否已确认"])
        writer.writeheader()
        writer.writerows([
            {"中文术语": "电阻", "英文翻译": "resistor", "领域": "电路", "添加日期": date.today().isoformat(), "是否已确认": "是"},
            {"中文术语": "电容", "英文翻译": "capacitor", "领域": "电路", "添加日期": date.today().isoformat(), "是否已确认": "是"},
            {"中文术语": "欧拉定律", "英文翻译": "Euler's Law", "领域": "电路", "添加日期": date.today().isoformat(), "是否已确认": "是"},
        ])
    return csv_path


@pytest.fixture
def test_pptx(tmp_path):
    """Create a minimal test PPTX file."""
    pptx_path = tmp_path / "test.pptx"

    slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="1" name="TextBox 1"/>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="0" y="0"/>
                        <a:ext cx="1000000" cy="500000"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst/>
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="zh-CN" dirty="0"/>
                            <a:t>电阻</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>'''

    import zipfile
    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("ppt/slides/slide1.xml", slide_xml)
        zf.writestr("[Content_Types].xml", content_types)

    return pptx_path
```

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "init: project structure"
```

---

## Task 2: config.py - 配置管理

**Files:**
- Create: `src/ppt_translator/config.py`
- Create: `tests/test_config.py`

- [ ] **Step 1: Write test for config**

```python
import os
from src.ppt_translator.config import Config, get_api_key, get_glossary_path

def test_config_loads_api_key(monkeypatch):
    monkeypatch.setenv("MINIMAX_API_KEY", "test-key-123")
    config = Config()
    assert config.api_key == "test-key-123"

def test_config_glossary_path():
    config = Config()
    assert "电路术语表.csv" in str(config.glossary_path)

def test_get_api_key_raises_without_key(monkeypatch):
    monkeypatch.delenv("MINIMAX_API_KEY", raising=False)
    try:
        get_api_key()
        assert False, "Should have raised SystemExit"
    except SystemExit:
        pass
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_config.py -v`
Expected: FAIL - ModuleNotFoundError

- [ ] **Step 3: Write minimal config.py**

```python
"""Configuration management for PPT Translator."""

import os
from pathlib import Path

DEFAULT_GLOSSARY = "电路术语表.csv"
GLOSSARY_DIR = Path(__file__).parent.parent


class Config:
    """Application configuration."""

    def __init__(self):
        self.api_key = os.environ.get("MINIMAX_API_KEY", "")
        self.glossary_path = GLOSSARY_DIR / DEFAULT_GLOSSARY
        self.minimax_model = "abab6.5s-chat"
        self.minimax_api_host = "https://api.minimax.chat/v1"
        self.request_timeout = 60
        self.default_font = "Arial"
        self.font_size_min = 8
        self.font_size_shrink_ratio = 0.5


def get_api_key() -> str:
    """Get MiniMax API key from environment."""
    api_key = os.environ.get("MINIMAX_API_KEY", "")
    if not api_key:
        print("Error: MINIMAX_API_KEY environment variable not set")
        print("Please set it with: export MINIMAX_API_KEY='your-key'")
        raise SystemExit(1)
    return api_key


def get_glossary_path() -> Path:
    """Get path to terminology glossary CSV."""
    return GLOSSARY_DIR / DEFAULT_GLOSSARY
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_config.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add config module with API key management"
```

---

## Task 3: terminology.py - 术语表管理

**Files:**
- Create: `src/ppt_translator/terminology.py`
- Create: `tests/test_terminology.py`

**Dependencies:** Task 2 (config.py)

- [ ] **Step 1: Write test for terminology**

```python
import csv
from pathlib import Path
from src.ppt_translator.terminology import TerminologyManager

def test_load_glossary(test_glossary_csv):
    """Test loading glossary from CSV."""
    manager = TerminologyManager(test_glossary_csv)
    assert len(manager.terms) > 0
    assert manager.get_english("电阻") == "resistor"
    assert manager.get_english("电容") == "capacitor"

def test_pre_replace_with_space():
    """Test terminology pre-replacement adds space only at word boundaries."""
    manager = TerminologyManager()
    manager.terms["电阻"] = "resistor"
    manager.terms["欧拉定律"] = "Euler's Law"

    # Test normal case
    text = "电阻R1"
    result = manager.pre_replace(text)
    assert "resistor " in result  # space after term

    # Test at end of sentence - no trailing space added
    text2 = "这是欧拉定律。"
    result2 = manager.pre_replace(text2)
    # The period should come right after the term, not after a space
    assert "Euler's Law." in result2

def test_add_new_term():
    """Test adding new term to glossary."""
    manager = TerminologyManager(test_glossary_csv)
    initial_count = len(manager.terms)
    manager.add_term("测试术语", "test term", "电路")
    assert len(manager.terms) == initial_count + 1
    assert manager.get_english("测试术语") == "test term"

def test_discover_term():
    """Test discovering new terms during translation."""
    manager = TerminologyManager()
    manager.discover_term("新术语", "new term")
    summary = manager.get_new_terms_summary()
    assert ("新术语", "new term") in summary
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_terminology.py -v`
Expected: FAIL - ModuleNotFoundError

- [ ] **Step 3: Write terminology.py**

```python
"""Terminology management for PPT translation."""

import csv
import re
from datetime import date
from pathlib import Path
from typing import Optional

from src.ppt_translator.config import get_glossary_path


class TerminologyManager:
    """Manages terminology glossary for translation."""

    def __init__(self, glossary_path: Optional[Path] = None):
        self.glossary_path = glossary_path or get_glossary_path()
        self.terms: dict[str, str] = {}
        self.new_terms: list[tuple[str, str]] = []  # (zh, en) discovered during translation
        self._load_glossary()

    def _load_glossary(self):
        """Load terminology from CSV file."""
        if not self.glossary_path.exists():
            return

        with open(self.glossary_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("是否已确认", "") == "是":
                    zh = row.get("中文术语", "").strip()
                    en = row.get("英文翻译", "").strip()
                    if zh and en:
                        self.terms[zh] = en

    def get_english(self, chinese: str) -> Optional[str]:
        """Get English translation for a Chinese term."""
        return self.terms.get(chinese)

    def pre_replace(self, text: str) -> str:
        """Pre-replace known terms in text before API translation.

        Adds space after replaced term only when followed by a word character.
        Preserves spacing at end of text and before punctuation.
        """
        result = text
        for zh, en in sorted(self.terms.items(), key=lambda x: len(x[0]), reverse=True):
            if zh in result:
                # Use regex to replace with proper spacing
                # Pattern: term followed by word char -> replace with term + space
                # Pattern: term at end -> replace without extra space
                # Pattern: term before punctuation -> replace without extra space
                pattern = re.escape(zh)
                # Replace term followed by alphanumeric or underscore
                result = re.sub(f'{pattern}(?=[a-zA-Z0-9_])', en + ' ', result)
                # Replace term at end of string
                result = re.sub(f'{pattern}$', en, result)
                # Replace term followed by punctuation
                result = re.sub(f'{pattern}(?=[.,;:!?，。；：！？])', en, result)
        return result

    def add_term(self, chinese: str, english: str, domain: str = "电路"):
        """Add a new term to the glossary."""
        self.terms[chinese] = english
        self.new_terms.append((chinese, english))

    def save_glossary(self):
        """Save all terms (including new ones) to CSV."""
        today = date.today().isoformat()

        # Read existing rows to preserve all fields
        existing_rows = []
        if self.glossary_path.exists():
            with open(self.glossary_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                existing_rows = list(reader)

        # Add new terms
        for zh, en in self.new_terms:
            existing_rows.append({
                "中文术语": zh,
                "英文翻译": en,
                "领域": "电路",
                "添加日期": today,
                "是否已确认": "是"
            })

        # Write back
        with open(self.glossary_path, "w", encoding="utf-8", newline="") as f:
            fieldnames = ["中文术语", "英文翻译", "领域", "添加日期", "是否已确认"]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(existing_rows)

        self.new_terms.clear()

    def discover_term(self, chinese: str, english: str):
        """Record a newly discovered term for later confirmation."""
        if chinese not in self.terms and (chinese, english) not in self.new_terms:
            self.new_terms.append((chinese, english))

    def get_new_terms_summary(self) -> list[tuple[str, str]]:
        """Get list of newly discovered terms."""
        return self.new_terms.copy()

    def clear_new_terms(self):
        """Clear pending new terms without saving."""
        self.new_terms.clear()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_terminology.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add terminology management module"
```

---

## Task 4: text_box.py - 文本框模型

**Files:**
- Create: `src/ppt_translator/text_box.py`
- Create: `tests/test_text_box.py`

**Dependencies:** Task 2 (config.py)

- [ ] **Step 1: Write test for text_box**

```python
from src.ppt_translator.text_box import TextBox, TextBoxCollection

def test_text_box_creation():
    """Test creating a text box."""
    tb = TextBox(
        shape_id="1",
        shape_name="TextBox 1",
        original_text="电阻",
        xpath="/p:sp/p:txBody/a:p/a:r/a:t"
    )
    assert tb.original_text == "电阻"
    assert tb.translated_text is None
    assert not tb.is_translated

def test_text_box_mark_translated():
    """Test marking text box as translated."""
    tb = TextBox(shape_id="1", shape_name="TextBox 1", original_text="电阻")
    tb.mark_translated("resistor")
    assert tb.translated_text == "resistor"
    assert tb.is_translated

def test_text_box_rollback():
    """Test rolling back translation."""
    tb = TextBox(shape_id="1", shape_name="TextBox 1", original_text="电阻")
    tb.mark_translated("resistor")
    tb.rollback()
    assert tb.translated_text is None
    assert not tb.is_translated
    assert tb.failed is False

def test_text_box_mark_failed():
    """Test marking translation as failed."""
    tb = TextBox(shape_id="1", shape_name="TextBox 1", original_text="电阻")
    tb.mark_failed("API timeout")
    assert tb.failed is True
    assert tb.error_message == "API timeout"

def test_collection():
    """Test text box collection."""
    coll = TextBoxCollection()
    tb1 = TextBox(shape_id="1", shape_name="TextBox 1", original_text="电阻")
    tb2 = TextBox(shape_id="2", shape_name="TextBox 2", original_text="电容")
    coll.add(tb1)
    coll.add(tb2)
    assert len(coll) == 2
    assert coll.get_by_id("1") == tb1
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_text_box.py -v`
Expected: FAIL - ModuleNotFoundError

- [ ] **Step 3: Write text_box.py**

```python
"""Text box model for tracking translation at text box level."""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TextBox:
    """Represents a text box (shape) in a PPTX slide.

    Each text box is treated as an independent translation unit.
    This allows for rollback at the text box level if translation fails.
    """

    shape_id: str
    shape_name: str
    original_text: str
    xpath: str  # XPath to the text element within the shape

    translated_text: Optional[str] = None
    failed: bool = False
    error_message: Optional[str] = None

    @property
    def is_translated(self) -> bool:
        return self.translated_text is not None

    def mark_translated(self, translated: str):
        """Mark this text box as successfully translated."""
        self.translated_text = translated
        self.failed = False
        self.error_message = None

    def mark_failed(self, error: str):
        """Mark this text box translation as failed."""
        self.failed = True
        self.error_message = error
        self.translated_text = None

    def rollback(self):
        """Rollback translation, reverting to original text."""
        self.translated_text = None
        self.failed = False
        self.error_message = None

    def get_final_text(self) -> str:
        """Get the final text to use (translated or original)."""
        if self.translated_text is not None and not self.failed:
            return self.translated_text
        return self.original_text


class TextBoxCollection:
    """Collection of text boxes from a PPTX."""

    def __init__(self):
        self._boxes: dict[str, TextBox] = {}

    def add(self, box: TextBox):
        """Add a text box to the collection."""
        self._boxes[box.shape_id] = box

    def get_by_id(self, shape_id: str) -> Optional[TextBox]:
        """Get a text box by its shape ID."""
        return self._boxes.get(shape_id)

    def __len__(self) -> int:
        return len(self._boxes)

    def __iter__(self):
        return iter(self._boxes.values())

    def get_failed(self) -> list[TextBox]:
        """Get all text boxes that failed translation."""
        return [box for box in self._boxes.values() if box.failed]

    def get_successful(self) -> list[TextBox]:
        """Get all successfully translated text boxes."""
        return [box for box in self._boxes.values() if box.is_translated and not box.failed]

    def summary(self) -> dict:
        """Get a summary of translation status."""
        total = len(self._boxes)
        succeeded = len(self.get_successful())
        failed = len(self.get_failed())
        pending = total - succeeded - failed

        return {
            "total": total,
            "succeeded": succeeded,
            "failed": failed,
            "pending": pending
        }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_text_box.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add text box model for granular tracking"
```

---

## Task 5: translator.py - MiniMax API调用

**Files:**
- Create: `src/ppt_translator/translator.py`
- Create: `tests/test_translator.py`

**Dependencies:** Task 2 (config.py)

- [ ] **Step 1: Write test for translator**

```python
import pytest
from unittest.mock import patch, MagicMock
from src.ppt_translator.translator import MiniMaxTranslator

def test_translate_text():
    """Test basic translation."""
    translator = MiniMaxTranslator(api_key="test-key")

    mock_response = MagicMock()
    mock_response.json.return_value = {
        "choices": [{
            "messages": [{
                "content": "Hello"
            }]
        }]
    }
    mock_response.raise_for_status = MagicMock()

    with patch.object(translator.session, "post", return_value=mock_response) as mock_post:
        result = translator.translate("你好")
        assert result == "Hello"
        mock_post.assert_called_once()

def test_translate_calls_api_with_correct_params():
    """Test API is called with correct parameters."""
    translator = MiniMaxTranslator(api_key="test-key")

    mock_response = MagicMock()
    mock_response.json.return_value = {"choices": [{"messages": [{"content": "test"}]}]}
    mock_response.raise_for_status = MagicMock()

    with patch.object(translator.session, "post", return_value=mock_response) as mock_post:
        translator.translate("测试")
        call_kwargs = mock_post.call_args[1]
        assert "MINIMAX_API_HOST" in call_kwargs["url"]
        assert call_kwargs["json"]["model"] == "abab6.5s-chat"

def test_translate_empty_text():
    """Test that empty text returns as-is."""
    translator = MiniMaxTranslator(api_key="test-key")
    result = translator.translate("")
    assert result == ""

def test_translate_timeout():
    """Test timeout handling."""
    import requests
    translator = MiniMaxTranslator(api_key="test-key")

    with patch.object(translator.session, "post", side_effect=requests.exceptions.Timeout()):
        try:
            translator.translate("test")
            assert False, "Should have raised"
        except TimeoutError as e:
            assert "timed out" in str(e)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_translator.py -v`
Expected: FAIL - ModuleNotFoundError

- [ ] **Step 3: Write translator.py**

```python
"""MiniMax API translator for PPT content."""

import re
import requests
from typing import Optional

from src.ppt_translator.config import get_api_key, Config


class MiniMaxTranslator:
    """Translates text using MiniMax API."""

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or get_api_key()
        self.config = Config()
        self.session = requests.Session()

    def translate(
        self,
        text: str,
        source_lang: str = "Chinese",
        target_lang: str = "English"
    ) -> str:
        """Translate text from source language to target language.

        Args:
            text: Text to translate
            source_lang: Source language name
            target_lang: Target language name

        Returns:
            Translated text
        """
        if not text or not text.strip():
            return text

        prompt = self._build_prompt(text, source_lang, target_lang)

        try:
            response = self.session.post(
                f"{self.config.minimax_api_host}/text/chatcompletion_v2",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": self.config.minimax_model,
                    "messages": [
                        {"role": "user", "content": prompt}
                    ]
                },
                timeout=self.config.request_timeout
            )
            response.raise_for_status()
            result = response.json()

            translated = result.get("choices", [{}])[0].get("messages", [{}])[0].get("content", "")
            return self._clean_translation(translated)

        except requests.exceptions.Timeout:
            raise TimeoutError(f"Translation request timed out after {self.config.request_timeout}s")
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"Translation API error: {e}")

    def _build_prompt(self, text: str, source_lang: str, target_lang: str) -> str:
        """Build translation prompt."""
        return f"""Translate the following {source_lang} text to {target_lang}.

Rules:
- Preserve technical terminology accurately
- Keep formulas and equations unchanged (e.g., E=mc², U=IR)
- Keep numbered lists and references as-is (e.g., 1-1, (a), (b), ① ② ③)
- Chinese person names should be converted to Pinyin (e.g., 欧拉 -> Oula)
- Use proper punctuation
- Keep the same tone and formality level

Text to translate:
{text}

{target_lang} translation:"""

    def _clean_translation(self, text: str) -> str:
        """Clean up translation output."""
        text = text.strip()
        # Remove any leading/trailing quotes if present
        text = re.sub(r'^["\'](.*)["\']$', r'\1', text)
        return text

    def translate_batch(self, texts: list[str]) -> list[str]:
        """Translate multiple texts.

        Args:
            texts: List of texts to translate

        Returns:
            List of translated texts
        """
        return [self.translate(t) for t in texts]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_translator.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add MiniMax API translator module"
```

---

## Task 6: ppt_handler.py - PPT处理（完整版）

**Files:**
- Create: `src/ppt_translator/ppt_handler.py`
- Create: `tests/test_ppt_handler.py`

**Dependencies:** Task 2 (config.py), Task 4 (text_box.py), Task 5 (translator.py)

**关键功能：**
1. 按文本框级别提取和替换文本
2. 设置翻译后文字的西文字体
3. 处理autoFit属性
4. 文本框级别回滚

- [ ] **Step 1: Write test for ppt_handler**

```python
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch, MagicMock
from src.ppt_translator.ppt_handler import PPTHandler, unpack_pptx, pack_pptx
from src.ppt_translator.text_box import TextBoxCollection

def test_unpack_pptx(tmp_path, test_pptx):
    """Test unpacking a PPTX file."""
    output_dir = tmp_path / "unpacked"
    unpack_pptx(test_pptx, output_dir)
    assert (output_dir / "ppt/slides/slide1.xml").exists()

def test_extract_text_boxes(test_pptx, tmp_path):
    """Test extracting text boxes at shape level."""
    output_dir = tmp_path / "unpacked"
    unpack_pptx(test_pptx, output_dir)

    handler = PPTHandler()
    handler.temp_dir = output_dir

    boxes = handler.extract_text_boxes()
    assert len(boxes) == 1
    box = boxes.get_by_id("1")
    assert box.original_text == "电阻"

def test_apply_translation_with_font(test_pptx, tmp_path):
    """Test applying translation and setting font."""
    output_dir = tmp_path / "unpacked"
    unpack_pptx(test_pptx, output_dir)

    handler = PPTHandler()
    handler.temp_dir = output_dir

    boxes = handler.extract_text_boxes()
    box = boxes.get_by_id("1")
    box.mark_translated("resistor")

    # Apply translations
    handler.apply_translations(boxes)

    # Read back the slide and check
    slide_path = output_dir / "ppt/slides/slide1.xml"
    content = slide_path.read_text(encoding="utf-8")

    # Check translated text is present
    assert "resistor" in content

    # Check font is set to Western font
    assert "Arial" in content or "Calibri" in content

def test_text_box_level_rollback(test_pptx, tmp_path):
    """Test rollback at text box level."""
    output_dir = tmp_path / "unpacked"
    unpack_pptx(test_pptx, output_dir)

    handler = PPTHandler()
    handler.temp_dir = output_dir

    boxes = handler.extract_text_boxes()
    box = boxes.get_by_id("1")
    box.mark_translated("resistor")

    # Apply then rollback
    handler.apply_translations(boxes)
    box.rollback()
    handler.apply_translations(boxes)

    # Original should still be there
    slide_path = output_dir / "ppt/slides/slide1.xml"
    content = slide_path.read_text(encoding="utf-8")
    assert "电阻" in content
    assert "resistor" not in content
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_ppt_handler.py -v`
Expected: FAIL - ModuleNotFoundError

- [ ] **Step 3: Write ppt_handler.py**

```python
"""PPT Handler - unpack, edit, and pack PPTX files with font and autoFit support."""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional
import re
import copy

from src.ppt_translator.config import Config
from src.ppt_translator.text_box import TextBox, TextBoxCollection


# Namespaces used in PPTX
NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Register namespaces to avoid ns0 prefixes
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class PPTHandler:
    """Handles PPTX file operations: unpack, edit, pack, font handling."""

    def __init__(self, config: Optional[Config] = None):
        self.config = config or Config()
        self.temp_dir: Optional[Path] = None
        self.pptx_path: Optional[Path] = None

    def unpack(self, pptx_path: Path, output_dir: Path):
        """Unpack PPTX to directory structure.

        Args:
            pptx_path: Path to input PPTX file
            output_dir: Directory to extract contents
        """
        self.pptx_path = pptx_path
        self.temp_dir = output_dir
        output_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(pptx_path, "r") as zf:
            zf.extractall(output_dir)

    def pack(self, output_path: Path):
        """Pack directory back to PPTX.

        Args:
            output_path: Path for output PPTX file
        """
        if not self.temp_dir:
            raise RuntimeError("No unpacked directory to pack")

        output_path.parent.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in self.temp_dir.rglob("*"):
                if file_path.is_file():
                    arcname = file_path.relative_to(self.temp_dir)
                    zf.write(file_path, arcname)

    def get_slide_files(self) -> list[Path]:
        """Get list of slide XML files."""
        if not self.temp_dir:
            return []

        slides_dir = self.temp_dir / "ppt" / "slides"
        if not slides_dir.exists():
            return []

        return sorted(slides_dir.glob("slide*.xml"))

    def extract_text_boxes(self) -> TextBoxCollection:
        """Extract text boxes from all slides.

        Returns:
            TextBoxCollection with all text boxes found
        """
        collection = TextBoxCollection()

        for slide_path in self.get_slide_files():
            content = slide_path.read_text(encoding="utf-8")
            slide_boxes = self._extract_text_boxes_from_slide(slide_path, content)
            for box in slide_boxes:
                collection.add(box)

        return collection

    def _extract_text_boxes_from_slide(self, slide_path: Path, xml_content: str) -> list[TextBox]:
        """Extract text boxes from a single slide.

        Returns list of TextBox objects with shape ID, name, and text content.
        """
        boxes = []
        root = ET.fromstring(xml_content)

        # Find all shape elements (sp)
        for sp_elem in root.iter():
            if sp_elem.tag.endswith("}sp"):
                # Get shape ID and name
                cNvPr = sp_elem.find(".//{*}cNvPr")
                if cNvPr is None:
                    continue

                shape_id = cNvPr.get("id", "")
                shape_name = cNvPr.get("name", "")

                # Skip if no text body
                txBody = sp_elem.find("{*}txBody")
                if txBody is None:
                    continue

                # Find all text runs
                text_parts = []
                for t_elem in tx_body_iter(txBody):
                    if t_elem.text:
                        text_parts.append(t_elem.text)

                if not text_parts:
                    continue

                original_text = "".join(text_parts)

                # Get the XPath to this text element (for replacement)
                # We'll use a simpler approach: find all a:t elements and build paths
                xpath = f"//{slide_path.name}//{shape_id}"

                box = TextBox(
                    shape_id=shape_id,
                    shape_name=shape_name,
                    original_text=original_text,
                    xpath=xpath
                )
                boxes.append(box)

        return boxes

    def apply_translations(self, boxes: TextBoxCollection):
        """Apply translations to slides and set fonts.

        For each text box that was successfully translated:
        1. Replace the text content
        2. Set the font to a Western font (Arial/Calibri)
        3. Enable autoFit if text length increased
        """
        for slide_path in self.get_slide_files():
            content = slide_path.read_text(encoding="utf-8")
            root = ET.fromstring(content)

            modified = False

            for sp_elem in root.iter():
                if sp_elem.tag.endswith("}sp"):
                    cNvPr = sp_elem.find(".//{*}cNvPr")
                    if cNvPr is None:
                        continue

                    shape_id = cNvPr.get("id", "")
                    box = boxes.get_by_id(shape_id)

                    if box is None or not box.is_translated or box.failed:
                        continue

                    # Find txBody
                    txBody = sp_elem.find("{*}txBody")
                    if txBody is None:
                        continue

                    # Replace text in all a:t elements within this shape
                    for t_elem in tx_body_iter(txBody):
                        if t_elem.text == box.original_text:
                            t_elem.text = box.get_final_text()
                            modified = True

                            # Set Western font on the run properties
                            self._set_western_font(txBody, self.config.default_font)

                            # Handle autoFit for character expansion
                            self._handle_autofit(txBody, box.original_text, box.get_final_text())

            if modified:
                # Write back
                slide_path.write_text(ET.tostring(root, encoding="unicode"), encoding="utf-8")

    def _set_western_font(self, txBody, font_name: str):
        """Set Western font on text runs in a text body."""
        for rPr in txBody.iter():
            if rPr.tag.endswith("}rPr"):
                # Find or create latinFont (ea for East Asian, mso for other)
                latin = rPr.find("{*}latin")
                if latin is None:
                    # Create new latin element
                    latin = ET.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
                    latin.set("typeface", font_name)

                    # Also set cs (complex script) font
                    cs = rPr.find("{*}cs")
                    if cs is None:
                        cs = ET.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}cs")
                    cs.set("typeface", font_name)

                    # And ea (East Asian) font - keep but ensure we have Western fallback
                    ea = rPr.find("{*}ea")
                    if ea is None:
                        ea = ET.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
                        ea.set("typeface", font_name)
                else:
                    latin.set("typeface", font_name)

    def _handle_autofit(self, txBody, original: str, translated: str):
        """Handle autoFit for character expansion.

        If translated text is longer than original, enable autoFit
        and set appropriate font size limits.
        """
        if len(translated) <= len(original):
            return

        # Calculate expansion ratio
        ratio = len(original) / len(translated) if len(translated) > 0 else 1.0

        # Find bodyPr element
        bodyPr = txBody.find("{*}bodyPr")
        if bodyPr is None:
            bodyPr = ET.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr")
            txBody.insert(0, bodyPr)

        # Enable autoFit
        bodyPr.set("fit", "text")

        # Calculate new font size if we need to shrink
        if ratio < self.config.font_size_shrink_ratio:
            # Find existing font size
            for rPr in txBody.iter():
                if rPr.tag.endswith("}rPr"):
                    sz = rPr.get("sz")
                    if sz:
                        new_sz = int(int(sz) * ratio)
                        # Enforce minimum
                        new_sz = max(new_sz, self.config.font_size_min * 100)  # PPT uses hundredths of pt
                        rPr.set("sz", str(new_sz))


def tx_body_iter(txBody):
    """Iterate over all a:t (text) elements in a text body."""
    for elem in txBody.iter():
        if elem.tag.endswith("}t"):
            yield elem


def unpack_pptx(pptx_path: Path, output_dir: Path):
    """Standalone function to unpack PPTX."""
    handler = PPTHandler()
    handler.unpack(pptx_path, output_dir)


def pack_pptx(input_dir: Path, output_path: Path):
    """Standalone function to pack PPTX."""
    handler = PPTHandler()
    handler.temp_dir = input_dir
    handler.pack(output_path)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_ppt_handler.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add PPT handler with font and autoFit support"
```

---

## Task 7: cli.py - 命令行接口（完整版）

**Files:**
- Create: `src/ppt_translator/cli.py`
- Create: `tests/test_cli.py`

**Dependencies:** Task 2, Task 3, Task 4, Task 5, Task 6

- [ ] **Step 1: Write test for cli**

```python
import pytest
from click.testing import CliRunner
from unittest.mock import patch, MagicMock
from src.ppt_translator.cli import translate, status

def test_translate_command_without_api_key(monkeypatch, tmp_path):
    """Test translate command fails without API key."""
    monkeypatch.delenv("MINIMAX_API_KEY", raising=False)
    runner = CliRunner()
    result = runner.invoke(translate, [str(tmp_path / "test.pptx")])
    assert result.exit_code != 0
    assert "MINIMAX_API_KEY" in result.output

def test_status_command(test_glossary_csv, monkeypatch):
    """Test status command shows glossary info."""
    monkeypatch.setenv("MINIMAX_API_KEY", "test-key")
    runner = CliRunner()
    result = runner.invoke(status, ["--glossary", str(test_glossary_csv)])
    assert result.exit_code == 0
    assert "Terms loaded" in result.output
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_cli.py -v`
Expected: FAIL - ModuleNotFoundError

- [ ] **Step 3: Write cli.py**

```python
"""CLI for PPT Translator."""

import sys
from pathlib import Path
from typing import Optional

import click

from src.ppt_translator.config import Config, get_api_key
from src.ppt_translator.terminology import TerminologyManager
from src.ppt_translator.translator import MiniMaxTranslator
from src.ppt_translator.ppt_handler import PPTHandler


@click.group()
def cli():
    """PPT Translator - Translate Chinese PPT to English."""
    pass


@cli.command()
@click.argument("input_pptx", type=click.Path(exists=True))
@click.option("--output-dir", "-o", type=click.Path(), help="Output directory")
@click.option("--glossary", "-g", type=click.Path(), help="Path to glossary CSV")
def translate(input_pptx: str, output_dir: Optional[str], glossary: Optional[str]):
    """Translate a Chinese PPTX file to English.

    INPUT_PPTX: Path to the input Chinese PPTX file
    """
    input_path = Path(input_pptx).resolve()

    # Setup output path
    if output_dir:
        output_path = Path(output_dir)
    else:
        translated_dir = input_path.parent / "Translated"
        translated_dir.mkdir(exist_ok=True)
        output_path = translated_dir / f"{input_path.stem}_英文版.pptx"

    # Setup glossary
    glossary_path = Path(glossary) if glossary else None
    term_manager = TerminologyManager(glossary_path)

    # Initialize translator
    try:
        api_key = get_api_key()
    except SystemExit:
        click.echo("Error: MINIMAX_API_KEY environment variable not set", err=True)
        sys.exit(1)

    translator = MiniMaxTranslator(api_key)

    # Process PPT
    click.echo(f"Processing: {input_path}")
    click.echo("Unpacking PPTX...")

    import tempfile
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        handler = PPTHandler()
        handler.unpack(input_path, temp_path)

        # Extract text boxes
        click.echo("Extracting text boxes...")
        boxes = handler.extract_text_boxes()

        click.echo(f"Found {len(boxes)} text boxes")

        # Translate each text box
        for box in boxes:
            if not box.original_text or not box.original_text.strip():
                continue

            click.echo(f"  Translating: {box.original_text[:30]}...")

            # Pre-replace known terms
            preprocessed = term_manager.pre_replace(box.original_text)

            # Translate via API
            try:
                translated = translator.translate(preprocessed)
                box.mark_translated(translated)

                # Check for new terms (heuristic: if translated contains different words)
                # This is simplified - real implementation would need smarter detection
                if box.original_text != translated:
                    # Record original->translated as potential new term
                    # We'll refine this in later iterations
                    pass

            except Exception as e:
                click.echo(f"    Warning: Translation failed: {e}")
                box.mark_failed(str(e))
                # Rollback - use original
                box.rollback()

        # Apply translations to slides
        click.echo("Applying translations...")
        handler.apply_translations(boxes)

        # Pack result
        click.echo(f"Packing to: {output_path}")
        handler.pack(output_path)

    # Show summary
    summary = boxes.summary()
    click.echo(f"\nTranslation complete!")
    click.echo(f"  Total text boxes: {summary['total']}")
    click.echo(f"  Succeeded: {summary['succeeded']}")
    click.echo(f"  Failed: {summary['failed']}")

    # Handle new terms discovery
    new_terms = term_manager.get_new_terms_summary()
    if new_terms:
        click.echo(f"\nDiscovered {len(new_terms)} new terms:")
        for zh, en in new_terms[:10]:
            click.echo(f"  {zh} -> {en}")
        if len(new_terms) > 10:
            click.echo(f"  ... and {len(new_terms) - 10} more")

        # Ask user if they want to add to glossary
        click.echo("\nAdd these terms to glossary? (y/n/a=add all/n=skip)")
        # For now, in non-interactive mode, just skip

    click.echo(f"\nOutput: {output_path}")


@cli.command()
@click.option("--glossary", "-g", type=click.Path(), help="Path to glossary CSV")
def status(glossary: Optional[str]):
    """Check terminology glossary status."""
    glossary_path = Path(glossary) if glossary else None
    term_manager = TerminologyManager(glossary_path)

    click.echo(f"Glossary: {term_manager.glossary_path}")
    click.echo(f"Terms loaded: {len(term_manager.terms)}")

    new_terms = term_manager.get_new_terms_summary()
    if new_terms:
        click.echo(f"Pending new terms: {len(new_terms)}")


if __name__ == "__main__":
    cli()
```

- [ ] **Step 4: Install dependencies and run tests**

```bash
pip install click pytest pytest-mock
pytest tests/test_cli.py -v
```

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add complete CLI interface with translate and status commands"
```

---

## Task 8: Skill文件创建

**Files:**
- Create: `skill.md` (Claude Skill definition)

**Dependencies:** Task 7 (cli.py)

- [ ] **Step 1: Create skill.md at correct location**

```markdown
# PPT Translator

Translate Chinese PPT files to English using MiniMax API with terminology management.

## Usage

```
/ppt-translator <path-to-pptx> [--output-dir <dir>] [--glossary <path>]
```

## Options

- `<path-to-pptx>`: Path to the Chinese PPTX file to translate
- `--output-dir`, `-o`: Output directory (default: source-dir/Translated/)
- `--glossary`, `-g`: Path to custom glossary CSV

## Examples

```
/ppt-translator 中文/第1章.pptx
/ppt-translator 中文/第2章.pptx --output-dir /path/to/output
```

## How It Works

1. Unpacks the PPTX file and extracts text boxes
2. Each text box is translated independently via MiniMax API
3. Pre-replaces known terminology from the glossary
4. Sets Western fonts (Arial/Calibri) on translated text
5. Enables autoFit for character expansion handling
6. Repacks the translated content into a new PPTX
7. Outputs to `<source-dir>/Translated/<filename>_英文版.pptx`

## Terminology

The tool maintains a terminology glossary at:
`~/.claude/skills/ppt-translator/电路术语表.csv`

The glossary contains electrical engineering terminology (373+ terms).
New terms discovered during translation can be added to the glossary.

## Translation Rules

- Technical terms are pre-replaced using the glossary before API translation
- Formulas (E=mc², U=IR) are preserved unchanged
- Numbered lists (1-1, (a), (b), ① ② ③) are preserved
- Chinese person names are converted to Pinyin

## Error Handling

- Text box level rollback: if a text box fails, it reverts to original
- Other text boxes continue processing
- Failed text boxes are reported in the summary

## Environment Variables

- `MINIMAX_API_KEY`: Your MiniMax API key (required)

## Requirements

- Python 3.10+
- click
- requests
```

Save to: `/Users/wusongfightstiger/.claude/skills/ppt-translator/skill.md`

- [ ] **Step 2: Commit**

```bash
git add -A
git commit -m "feat: add Claude Skill definition"
```

---

## Task 9: README文档

**Files:**
- Create: `README.md`

- [ ] **Step 1: Write README.md**

```markdown
# PPT Translator

Translate Chinese PPT files to English with terminology management.

## Quick Start

1. Set your MiniMax API key:
```bash
export MINIMAX_API_KEY='your-api-key'
```

2. Translate a PPT file:
```bash
python -m src.ppt_translator.cli translate 中文/第1章.pptx
```

3. Or use the Claude Skill:
```
/ppt-translator 中文/第1章.pptx
```

## Project Structure

```
ppt-translator/
├── src/ppt_translator/
│   ├── cli.py           # CLI interface
│   ├── config.py        # Configuration
│   ├── terminology.py    # Terminology management
│   ├── text_box.py      # Text box model
│   ├── translator.py     # MiniMax API
│   └── ppt_handler.py    # PPTX operations
├── tests/               # Tests
├── 电路术语表.csv       # Terminology glossary (373+ terms)
└── docs/               # Documentation
```

## Terminology Glossary

The glossary file `电路术语表.csv` contains pre-loaded electrical engineering terminology.

Format:
```csv
中文术语,英文翻译,领域,添加日期,是否已确认
电阻,resistor,电路,2026-03-21,是
```

## Development

```bash
# Install dependencies
pip install click requests pytest pytest-mock

# Run tests
pytest tests/ -v

# Format code
black src/
```

## Translation Process

1. **Unpack**: PPTX is extracted to temporary directory
2. **Extract Text Boxes**: Each shape with text is extracted as a TextBox
3. **Pre-replace**: Known terms from glossary are replaced before API call
4. **Translate**: Each text box is translated via MiniMax API
5. **Apply**: Translated text is inserted, fonts set to Arial/Calibri
6. **Pack**: Result is packaged as new PPTX

## Error Handling

- Text box level: if one text box fails, only that box is rolled back
- Batch level: other text boxes continue processing
- Summary reports succeeded/failed counts

## License

MIT
```

- [ ] **Step 2: Commit**

```bash
git add -A
git commit -m "docs: add README"
```

---

## Task 10: 集成测试

**Files:**
- Create: `tests/test_integration.py`

**Dependencies:** All previous tasks

- [ ] **Step 1: Write integration test**

```python
import pytest
import zipfile
from pathlib import Path
from unittest.mock import patch, MagicMock

from src.ppt_translator.cli import translate
from src.ppt_translator.terminology import TerminologyManager
from src.ppt_translator.translator import MiniMaxTranslator
from src.ppt_translator.ppt_handler import PPTHandler
from click.testing import CliRunner


def test_full_translation_flow(tmp_path, monkeypatch, test_glossary_csv):
    """Test complete translation workflow."""
    # Create test PPTX
    test_pptx = tmp_path / "test.pptx"

    slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="1" name="TextBox 1"/>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="zh-CN" dirty="0"/>
                            <a:t>电阻</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>'''

    with zipfile.ZipFile(test_pptx, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("ppt/slides/slide1.xml", slide_xml)
        zf.writestr("[Content_Types].xml", content_types)

    # Mock the API response
    def mock_translate(self, text, *args, **kwargs):
        if "resistor" in text.lower():
            return "resistor"
        return "translated"

    monkeypatch.setenv("MINIMAX_API_KEY", "test-key")

    with patch.object(MiniMaxTranslator, "translate", mock_translate):
        runner = CliRunner()
        result = runner.invoke(translate, [str(test_pptx), "--glossary", str(test_glossary_csv)])

        assert result.exit_code == 0, f"Failed: {result.output}"
        assert "Translation complete" in result.output
        assert "Succeeded: 1" in result.output

        # Check output file exists
        output_path = test_pptx.parent / "Translated" / "test_英文版.pptx"
        assert output_path.exists(), f"Output not found: {output_path}"
```

- [ ] **Step 2: Run integration test**

Run: `pytest tests/test_integration.py -v`
Expected: Should pass with mocks

- [ ] **Step 3: Commit**

```bash
git add -A
git commit -m "test: add integration test"
```

---

## 验证清单

After all tasks complete, verify:

1. [ ] `python -m src.ppt_translator.cli --help` works
2. [ ] `pytest tests/ -v` all tests pass
3. [ ] Terminology CSV is properly formatted and loadable
4. [ ] Skill file is in correct location
5. [ ] README accurately reflects project state
6. [ ] Test PPTX can be translated with fonts properly set
7. [ ] Text box level rollback works when translation fails
```

