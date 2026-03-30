"""Terminology management for PPT Translator."""

import csv
import re
from datetime import date
from pathlib import Path
from typing import Optional


class TerminologyManager:
    """Manager for terminology glossary."""

    def __init__(self, glossary_path: Path):
        """Initialize terminology manager.

        Args:
            glossary_path: Path to the glossary CSV file.
        """
        self.glossary_path = Path(glossary_path)
        self._terms: dict[str, str] = {}  # chinese -> english
        self._domains: dict[str, str] = {}  # chinese -> domain
        self._new_terms: list[dict] = []  # newly discovered terms
        self._load_glossary()

    def _load_glossary(self) -> None:
        """Load terminology from CSV file.

        Only loads terms where "是否已确认" is "是".
        """
        if not self.glossary_path.exists():
            return

        with open(self.glossary_path, "r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("是否已确认", "").strip() == "是":
                    chinese = row["中文术语"].strip()
                    english = row["英文翻译"].strip()
                    domain = row.get("领域", "").strip()
                    self._terms[chinese] = english
                    self._domains[chinese] = domain

    def get_english(self, chinese: str) -> Optional[str]:
        """Get English translation for a Chinese term.

        Args:
            chinese: Chinese terminology.

        Returns:
            English translation or None if not found.
        """
        return self._terms.get(chinese)

    def pre_replace(self, text: str) -> str:
        """Pre-replace Chinese terms with English translations in text.

        Adds space before English translation when the term is followed by
        alphanumeric characters, subscript markers, or certain punctuation
        that indicates the term continues.

        Args:
            text: Input text with Chinese terms.

        Returns:
            Text with terms replaced and proper spacing.
        """
        result = text

        # Sort terms by length (longest first) to handle overlapping terms
        sorted_terms = sorted(self._terms.items(), key=lambda x: len(x[0]), reverse=True)

        for chinese, english in sorted_terms:
            # Pattern: Chinese term followed by a boundary character
            # Boundary chars: alphanumeric, subscript numbers, some punctuation
            # We add a space before the English if followed by alphanumeric or subscript

            # Escape special regex characters in Chinese term
            escaped_chinese = re.escape(chinese)

            # Pattern to match Chinese term followed by a word character (alphanumeric)
            # or subscript marker (parsed as separate characters in some contexts)
            # We look for the term and check what follows
            pattern = re.compile(escaped_chinese + r'(?=[a-zA-Z0-9\u00B2\u00B3\u2070-\u2079])')

            if pattern.search(result):
                # Replace with English + space before it if followed by alphanumeric
                result = pattern.sub(english + " ", result)
            else:
                # Just do simple replacement
                result = result.replace(chinese, english)

        return result

    def add_term(self, chinese: str, english: str, domain: str = "通用") -> None:
        """Add a new terminology entry.

        Args:
            chinese: Chinese terminology.
            english: English translation.
            domain: Domain/category.
        """
        self._terms[chinese] = english
        self._domains[chinese] = domain

    def save_glossary(self) -> None:
        """Save the current glossary to CSV file.

        Writes all terms including newly added ones.
        """
        fieldnames = ["中文术语", "英文翻译", "领域", "添加日期", "是否已确认"]

        # Read existing data to preserve non-confirmed terms
        existing_terms = {}
        if self.glossary_path.exists():
            with open(self.glossary_path, "r", encoding="utf-8", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get("是否已确认", "").strip() != "是":
                        chinese = row["中文术语"].strip()
                        existing_terms[chinese] = row

        # Write all terms
        with open(self.glossary_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

            # Write confirmed terms
            for chinese, english in self._terms.items():
                writer.writerow({
                    "中文术语": chinese,
                    "英文翻译": english,
                    "领域": self._domains.get(chinese, "通用"),
                    "添加日期": date.today().isoformat(),
                    "是否已确认": "是"
                })

            # Write non-confirmed terms
            for term_data in existing_terms.values():
                writer.writerow(term_data)

    def discover_term(self, chinese: str, english: str) -> None:
        """Record a newly discovered term pair for later review.

        Args:
            chinese: Chinese terminology.
            english: English translation.
        """
        self._new_terms.append({
            "中文术语": chinese,
            "英文翻译": english,
            "领域": "待定",
            "添加日期": date.today().isoformat(),
            "是否已确认": "否"
        })

    def get_new_terms_summary(self) -> list[dict]:
        """Get list of newly discovered terms.

        Returns:
            List of dictionaries containing new term information.
        """
        return self._new_terms.copy()

    def clear_new_terms(self) -> None:
        """Clear all undiscovered saved new terms."""
        self._new_terms.clear()
