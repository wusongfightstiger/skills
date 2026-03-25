"""Tests for terminology module."""

import pytest
from pathlib import Path
from datetime import date

from ppt_translator.terminology import TerminologyManager


class TestTerminologyManager:
    """Test cases for TerminologyManager class."""

    def test_load_glossary(self, test_glossary_csv):
        """Test loading glossary from CSV.

        Confirms that '电阻' maps to 'resistor'.
        """
        manager = TerminologyManager(test_glossary_csv)

        # Verify key terms are loaded
        assert manager.get_english("电阻") == "resistor"
        assert manager.get_english("电容") == "capacitor"
        assert manager.get_english("欧拉定律") == "Euler's Law"

    def test_pre_replace_with_space(self, test_glossary_csv):
        """Test that pre_replace adds space when term is followed by alphanumeric."""
        manager = TerminologyManager(test_glossary_csv)

        # When term is followed by alphanumeric, should add space
        text = "电阻R1是重要的元件"
        result = manager.pre_replace(text)
        assert result == "resistor R1是重要的元件"

        # When term is followed by subscript-like characters
        text2 = "电阻²"
        result2 = manager.pre_replace(text2)
        assert result2 == "resistor ²" or result2 == "resistor²"

        # When term is followed by punctuation, no extra space
        text3 = "电阻。"
        result3 = manager.pre_replace(text3)
        assert result3 == "resistor。" or result3 == "resistor。"

        # When term is at end of string, no extra space
        text4 = "这是电阻"
        result4 = manager.pre_replace(text4)
        assert result4 == "这是resistor"

    def test_add_term(self, test_glossary_csv):
        """Test adding a new term and retrieving it."""
        manager = TerminologyManager(test_glossary_csv)

        # Add a new term
        manager.add_term("电感", "inductor", "电路")

        # Verify it can be retrieved
        assert manager.get_english("电感") == "inductor"

    def test_discover_term(self, test_glossary_csv):
        """Test discovering new terms."""
        manager = TerminologyManager(test_glossary_csv)

        # Discover new terms
        manager.discover_term("新型元件", "new component")
        manager.discover_term("高级电路", "advanced circuit")

        # Get summary
        summary = manager.get_new_terms_summary()

        assert len(summary) == 2

        # Check first discovered term
        term1 = summary[0]
        assert term1["中文术语"] == "新型元件"
        assert term1["英文翻译"] == "new component"
        assert term1["领域"] == "待定"
        assert term1["是否已确认"] == "否"
        assert "添加日期" in term1

        # Check second discovered term
        term2 = summary[1]
        assert term2["中文术语"] == "高级电路"
        assert term2["英文翻译"] == "advanced circuit"
        assert term2["领域"] == "待定"
        assert term2["是否已确认"] == "否"

    def test_clear_new_terms(self, test_glossary_csv):
        """Test clearing undiscovered new terms."""
        manager = TerminologyManager(test_glossary_csv)

        # Add some new terms
        manager.discover_term("待定术语", "pending term")
        manager.discover_term("另一个待定", "another pending")

        # Verify they exist
        assert len(manager.get_new_terms_summary()) == 2

        # Clear them
        manager.clear_new_terms()

        # Verify they are cleared
        assert len(manager.get_new_terms_summary()) == 0
