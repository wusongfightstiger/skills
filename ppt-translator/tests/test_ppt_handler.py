"""Tests for ppt_handler module."""

import pytest
from pathlib import Path

from ppt_translator.config import Config
from ppt_translator.ppt_handler import (
    PPTHandler,
    TextBox,
    TextBoxCollection,
    unpack_pptx,
    pack_pptx,
    tx_body_iter,
)


class TestTextBoxCollection:
    """Tests for TextBoxCollection class."""

    def test_add_and_iterate(self):
        """Test adding boxes and iterating."""
        collection = TextBoxCollection()
        box1 = TextBox(Path("slide1.xml"), 1, "TextBox 1", "Hello")
        box2 = TextBox(Path("slide1.xml"), 2, "TextBox 2", "World")

        collection.add(box1)
        collection.add(box2)

        assert len(collection) == 2
        assert list(collection) == [box1, box2]

    def test_get_by_slide(self):
        """Test getting boxes by slide path."""
        collection = TextBoxCollection()
        box1 = TextBox(Path("slide1.xml"), 1, "Box 1", "Text 1")
        box2 = TextBox(Path("slide2.xml"), 2, "Box 2", "Text 2")
        box3 = TextBox(Path("slide1.xml"), 3, "Box 3", "Text 3")

        collection.add(box1)
        collection.add(box2)
        collection.add(box3)

        slide1_boxes = collection.get_by_slide(Path("slide1.xml"))
        assert len(slide1_boxes) == 2
        assert box1 in slide1_boxes
        assert box3 in slide1_boxes

    def test_rollback(self):
        """Test rollback restores original text."""
        collection = TextBoxCollection()
        box = TextBox(Path("slide1.xml"), 1, "Box 1", "Original")
        box.text = "Translated"

        collection.add(box)
        collection.rollback()

        assert box.text == "Original"


class TestUnpackPack:
    """Tests for unpack and pack functions."""

    def test_unpack_pptx(self, test_pptx, tmp_path):
        """Test unpacking a PPTX file."""
        output_dir = tmp_path / "unpacked"
        unpack_pptx(test_pptx, output_dir)

        assert output_dir.exists()
        assert (output_dir / "ppt" / "slides" / "slide1.xml").exists()
        assert (output_dir / "[Content_Types].xml").exists()

    def test_pack_pptx(self, test_pptx, tmp_path):
        """Test packing a directory into PPTX."""
        output_dir = tmp_path / "unpacked"
        unpack_pptx(test_pptx, output_dir)

        repacked_path = tmp_path / "repacked.pptx"
        pack_pptx(repacked_path, output_dir)

        assert repacked_path.exists()


class TestPPTHandler:
    """Tests for PPTHandler class."""

    def test_unpack_pptx(self, test_pptx, tmp_path, mock_config):
        """Test PPTHandler unpack method."""
        handler = PPTHandler(mock_config)

        output_dir = tmp_path / "unpacked"
        handler.unpack(test_pptx, output_dir)

        assert handler._unpacked_dir == output_dir
        assert handler._slides_dir == output_dir / "ppt" / "slides"

    def test_get_slide_files(self, test_pptx, tmp_path, mock_config):
        """Test getting slide files list."""
        handler = PPTHandler(mock_config)

        output_dir = tmp_path / "unpacked"
        handler.unpack(test_pptx, output_dir)

        slide_files = handler.get_slide_files()
        assert len(slide_files) == 1
        assert slide_files[0].name == "slide1.xml"

    def test_extract_text_boxes(self, test_pptx, tmp_path, mock_config):
        """Test extracting text boxes from PPTX."""
        handler = PPTHandler(mock_config)

        output_dir = tmp_path / "unpacked"
        handler.unpack(test_pptx, output_dir)

        boxes = handler.extract_text_boxes()

        assert len(boxes) == 1
        box = list(boxes)[0]
        assert box.shape_name == "TextBox 1"
        assert box.text == "电阻"

    def test_apply_translation_with_font(self, test_pptx, tmp_path, mock_config):
        """Test applying translation and setting fonts."""
        handler = PPTHandler(mock_config)

        output_dir = tmp_path / "unpacked"
        handler.unpack(test_pptx, output_dir)

        # Extract and modify text
        boxes = handler.extract_text_boxes()
        box = list(boxes)[0]
        box.text = "Resistor"

        # Apply translations
        handler.apply_translations(boxes)

        # Verify the slide was modified
        import xml.etree.ElementTree as ET
        slide_path = handler.get_slide_files()[0]
        tree = ET.parse(slide_path)
        root = tree.getroot()

        # Check that fonts were set
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        found_latin = False
        for rPr in root.iter():
            if rPr.tag.endswith("}rPr"):
                latin = rPr.find("a:latin", ns)
                if latin is not None:
                    assert latin.get("typeface") == "Arial"
                    found_latin = True
                cs = rPr.find("a:cs", ns)
                if cs is not None:
                    assert cs.get("typeface") == "Arial"
                ea = rPr.find("a:ea", ns)
                if ea is not None:
                    assert ea.get("typeface") == "Arial"

        assert found_latin, "Should have found at least one latin font setting"

    def test_text_box_level_rollback(self, test_pptx, tmp_path, mock_config):
        """Test text box level rollback functionality."""
        handler = PPTHandler(mock_config)

        output_dir = tmp_path / "unpacked"
        handler.unpack(test_pptx, output_dir)

        # Extract text boxes
        boxes = handler.extract_text_boxes()
        original_box = list(boxes)[0]
        original_text = original_box.text

        # Modify the text
        original_box.text = "Modified Text"

        # Rollback
        boxes.rollback()

        # Verify rollback
        assert original_box.text == original_text

    def test_pack_after_unpack(self, test_pptx, tmp_path, mock_config):
        """Test pack creates valid PPTX after unpack."""
        handler = PPTHandler(mock_config)

        output_dir = tmp_path / "unpacked"
        handler.unpack(test_pptx, output_dir)

        # Extract and modify
        boxes = handler.extract_text_boxes()
        box = list(boxes)[0]
        box.text = "Translated"

        handler.apply_translations(boxes)

        # Pack
        output_pptx = tmp_path / "output.pptx"
        handler.pack(output_pptx)

        assert output_pptx.exists()

        # Verify packed file can be reopened
        handler2 = PPTHandler(mock_config)
        handler2.unpack(output_pptx, tmp_path / "reopened")
        slide_files = handler2.get_slide_files()
        assert len(slide_files) == 1


class TestTxBodyIter:
    """Tests for tx_body_iter helper function."""

    def test_tx_body_iter(self):
        """Test iterating over txBody elements."""
        import xml.etree.ElementTree as ET

        xml_str = '''<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:sp>
                <p:txBody>
                    <a:p><a:r><a:t>Text 1</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:txBody>
                    <a:p><a:r><a:t>Text 2</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
        </root>'''

        root = ET.fromstring(xml_str)
        tx_bodies = list(tx_body_iter(root))

        assert len(tx_bodies) == 2
