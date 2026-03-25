"""Tests for text_box module."""

import pytest

from ppt_translator.text_box import TextBox, TextBoxCollection


class TestTextBox:
    """Test cases for TextBox dataclass."""

    def test_text_box_creation(self):
        """Test creating a TextBox with initial state."""
        tb = TextBox(
            shape_id=1,
            shape_name="TextBox 1",
            original_text="Hello",
            xpath="//p:sp[1]/p:txBody/a:p[1]/a:r[1]/a:t"
        )

        assert tb.shape_id == 1
        assert tb.shape_name == "TextBox 1"
        assert tb.original_text == "Hello"
        assert tb.xpath == "//p:sp[1]/p:txBody/a:p[1]/a:r[1]/a:t"
        assert tb.translated_text is None
        assert tb.failed is False
        assert tb.error_message is None

    def test_text_box_mark_translated(self):
        """Test marking a TextBox as translated."""
        tb = TextBox(
            shape_id=1,
            shape_name="TextBox 1",
            original_text="Hello",
            xpath="//p:sp[1]"
        )

        assert tb.is_translated() is False

        tb.mark_translated("你好")

        assert tb.is_translated() is True
        assert tb.translated_text == "你好"
        assert tb.failed is False
        assert tb.error_message is None

    def test_text_box_rollback(self):
        """Test rolling back a translation."""
        tb = TextBox(
            shape_id=1,
            shape_name="TextBox 1",
            original_text="Hello",
            xpath="//p:sp[1]"
        )

        tb.mark_translated("你好")
        assert tb.is_translated() is True

        tb.rollback()

        assert tb.is_translated() is False
        assert tb.translated_text is None
        assert tb.failed is False
        assert tb.error_message is None

    def test_text_box_mark_failed(self):
        """Test marking a TextBox as failed."""
        tb = TextBox(
            shape_id=1,
            shape_name="TextBox 1",
            original_text="Hello",
            xpath="//p:sp[1]"
        )

        tb.mark_failed("API rate limit exceeded")

        assert tb.is_translated() is False
        assert tb.failed is True
        assert tb.error_message == "API rate limit exceeded"
        assert tb.translated_text is None

    def test_text_box_get_final_text_original(self):
        """Test get_final_text returns original when not translated."""
        tb = TextBox(
            shape_id=1,
            shape_name="TextBox 1",
            original_text="Hello",
            xpath="//p:sp[1]"
        )

        assert tb.get_final_text() == "Hello"

    def test_text_box_get_final_text_translated(self):
        """Test get_final_text returns translated text when available."""
        tb = TextBox(
            shape_id=1,
            shape_name="TextBox 1",
            original_text="Hello",
            xpath="//p:sp[1]"
        )

        tb.mark_translated("你好")

        assert tb.get_final_text() == "你好"


class TestTextBoxCollection:
    """Test cases for TextBoxCollection class."""

    def test_collection_add_and_get(self):
        """Test adding and retrieving text boxes."""
        collection = TextBoxCollection()

        tb1 = TextBox(shape_id=1, shape_name="Box 1", original_text="Hello", xpath="//sp[1]")
        tb2 = TextBox(shape_id=2, shape_name="Box 2", original_text="World", xpath="//sp[2]")

        collection.add(tb1)
        collection.add(tb2)

        assert len(collection) == 2
        assert collection.get_by_id(1) is tb1
        assert collection.get_by_id(2) is tb2
        assert collection.get_by_id(999) is None

    def test_collection_iter(self):
        """Test iterating over collection."""
        collection = TextBoxCollection()

        tb1 = TextBox(shape_id=1, shape_name="Box 1", original_text="Hello", xpath="//sp[1]")
        tb2 = TextBox(shape_id=2, shape_name="Box 2", original_text="World", xpath="//sp[2]")

        collection.add(tb1)
        collection.add(tb2)

        boxes = list(collection)
        assert len(boxes) == 2
        assert tb1 in boxes
        assert tb2 in boxes

    def test_collection_get_failed(self):
        """Test getting failed text boxes."""
        collection = TextBoxCollection()

        tb1 = TextBox(shape_id=1, shape_name="Box 1", original_text="Hello", xpath="//sp[1]")
        tb2 = TextBox(shape_id=2, shape_name="Box 2", original_text="World", xpath="//sp[2]")
        tb3 = TextBox(shape_id=3, shape_name="Box 3", original_text="Test", xpath="//sp[3]")

        tb1.mark_translated("你好")
        tb2.mark_failed("Error")
        tb3.mark_translated("测试")

        collection.add(tb1)
        collection.add(tb2)
        collection.add(tb3)

        failed = collection.get_failed()
        assert len(failed) == 1
        assert failed[0] is tb2

    def test_collection_get_successful(self):
        """Test getting successfully translated text boxes."""
        collection = TextBoxCollection()

        tb1 = TextBox(shape_id=1, shape_name="Box 1", original_text="Hello", xpath="//sp[1]")
        tb2 = TextBox(shape_id=2, shape_name="Box 2", original_text="World", xpath="//sp[2]")
        tb3 = TextBox(shape_id=3, shape_name="Box 3", original_text="Test", xpath="//sp[3]")

        tb1.mark_translated("你好")
        tb2.mark_translated("世界")
        tb3.mark_translated("测试")

        collection.add(tb1)
        collection.add(tb2)
        collection.add(tb3)

        successful = collection.get_successful()
        assert len(successful) == 3

    def test_collection_summary(self):
        """Test collection summary."""
        collection = TextBoxCollection()

        tb1 = TextBox(shape_id=1, shape_name="Box 1", original_text="Hello", xpath="//sp[1]")
        tb2 = TextBox(shape_id=2, shape_name="Box 2", original_text="World", xpath="//sp[2]")
        tb3 = TextBox(shape_id=3, shape_name="Box 3", original_text="Test", xpath="//sp[3]")

        tb1.mark_translated("你好")
        tb2.mark_failed("Error")
        # tb3 is pending

        collection.add(tb1)
        collection.add(tb2)
        collection.add(tb3)

        summary = collection.summary()
        assert summary == {"total": 3, "translated": 1, "failed": 1, "pending": 1}