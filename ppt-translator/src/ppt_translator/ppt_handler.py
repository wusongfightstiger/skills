"""PPT handling for PPT Translator.

This module provides functionality for unpacking, extracting text, and repacking
PowerPoint files (PPTX format).
"""

import xml.etree.ElementTree as ET
import zipfile
import shutil
from pathlib import Path
from typing import Iterator, Optional

from .config import Config


class TextBox:
    """Represents a text box (shape) in a PowerPoint slide.

    Attributes:
        slide_path: Path to the slide XML file.
        shape_id: The shape ID from cNvPr.
        shape_name: The shape name from cNvPr.
        text: The text content of the text box.
        original_text: The original text before translation.
    """

    def __init__(self, slide_path: Path, shape_id: int, shape_name: str, text: str):
        """Initialize a TextBox.

        Args:
            slide_path: Path to the slide XML file.
            shape_id: The shape ID.
            shape_name: The shape name.
            text: The text content.
        """
        self.slide_path = slide_path
        self.shape_id = shape_id
        self.shape_name = shape_name
        self.text = text
        self.original_text = text

    def __repr__(self) -> str:
        return f"TextBox(shape_id={self.shape_id}, name={self.shape_name}, text={self.text[:20]}...)"


class TextBoxCollection:
    """Collection of TextBox objects.

    Provides dictionary-like access by shape_id and slide_path.
    """

    def __init__(self):
        """Initialize an empty TextBoxCollection."""
        self._boxes: list[TextBox] = []

    def add(self, box: TextBox) -> None:
        """Add a TextBox to the collection.

        Args:
            box: The TextBox to add.
        """
        self._boxes.append(box)

    def __iter__(self) -> Iterator[TextBox]:
        """Iterate over all text boxes."""
        return iter(self._boxes)

    def __len__(self) -> int:
        """Return the number of text boxes."""
        return len(self._boxes)

    def get_by_slide(self, slide_path: Path) -> list[TextBox]:
        """Get all text boxes from a specific slide.

        Args:
            slide_path: Path to the slide file.

        Returns:
            List of TextBox objects from that slide.
        """
        return [box for box in self._boxes if box.slide_path == slide_path]

    def rollback(self) -> None:
        """Rollback all text boxes to their original text.

        This is useful when translation fails and we need to restore original text.
        """
        for box in self._boxes:
            box.text = box.original_text


# XML namespaces
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def unpack_pptx(pptx_path: Path, output_dir: Path) -> None:
    """Unpack a PPTX file to a directory.

    PPTX files are ZIP archives containing XML files. This function
    extracts all contents to the specified output directory.

    Args:
        pptx_path: Path to the PPTX file.
        output_dir: Directory to extract contents to.
    """
    output_dir = Path(output_dir)
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)

    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(output_dir)


def pack_pptx(output_path: Path, source_dir: Path) -> None:
    """Pack a directory into a PPTX file.

    Args:
        output_path: Path for the output PPTX file.
        source_dir: Directory containing the extracted PPTX contents.
    """
    output_path = Path(output_path)
    source_dir = Path(source_dir)

    if output_path.exists():
        output_path.unlink()

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in source_dir.rglob("*"):
            if file_path.is_file():
                arcname = file_path.relative_to(source_dir)
                zf.write(file_path, arcname)


def tx_body_iter(tree) -> Iterator:
    """Iterate over text bodies (txBody) in a slide XML tree.

    Args:
        tree: XML element tree or element.

    Yields:
        txBody elements found in the tree.
    """
    # Find all txBody elements regardless of namespace
    for elem in tree.iter():
        if elem.tag.endswith("}txBody") or elem.tag == "txBody":
            yield elem


def tx_run_iter(tree) -> Iterator:
    """Iterate over text runs (r) in a slide XML tree.

    Args:
        tree: XML element tree or element.

    Yields:
        r (run) elements found in txBody elements.
    """
    for txBody in tx_body_iter(tree):
        for elem in txBody:
            if elem.tag.endswith("}r") or elem.tag == "r":
                yield elem


def text_iter(tree) -> Iterator:
    """Iterate over text elements (t) in a slide XML tree.

    Args:
        tree: XML element tree or element.

    Yields:
        t (text) elements found in runs.
    """
    for run in tx_run_iter(tree):
        for elem in run:
            if elem.tag.endswith("}t") or elem.tag == "t":
                yield elem


class PPTHandler:
    """Handler for PowerPoint file operations.

    Provides methods for unpacking, extracting text, applying translations,
    and repacking PPTX files.
    """

    def __init__(self, config: Config):
        """Initialize the PPTHandler.

        Args:
            config: Configuration object with settings like default_font.
        """
        self.config = config
        self._unpacked_dir: Optional[Path] = None
        self._slides_dir: Optional[Path] = None

    def unpack(self, pptx_path: Path, output_dir: Path) -> None:
        """Unpack a PPTX file to a directory.

        Args:
            pptx_path: Path to the PPTX file.
            output_dir: Directory to extract contents to.
        """
        unpack_pptx(pptx_path, output_dir)
        self._unpacked_dir = Path(output_dir)
        self._slides_dir = self._unpacked_dir / "ppt" / "slides"

    def pack(self, output_path: Path) -> None:
        """Pack the unpacked contents into a PPTX file.

        Args:
            output_path: Path for the output PPTX file.
        """
        if self._unpacked_dir is None:
            raise ValueError("No unpacked directory set. Call unpack() first.")
        pack_pptx(output_path, self._unpacked_dir)

    def get_slide_files(self) -> list[Path]:
        """Get list of slide XML files.

        Returns:
            List of paths to slide XML files, sorted by slide number.
        """
        if self._slides_dir is None:
            raise ValueError("No slides directory set. Call unpack() first.")

        slide_files = sorted(
            [f for f in self._slides_dir.iterdir() if f.name.startswith("slide") and f.suffix == ".xml"],
            key=lambda x: int(x.stem.replace("slide", ""))
        )
        return slide_files

    def _extract_text_boxes_from_slide(self, slide_path: Path) -> list[TextBox]:
        """Extract text boxes from a single slide.

        Args:
            slide_path: Path to the slide XML file.

        Returns:
            List of TextBox objects found in the slide.
        """
        import xml.etree.ElementTree as ET

        tree = ET.parse(slide_path)
        root = tree.getroot()

        text_boxes = []

        # Define namespace URIs
        NS_P_URI = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"

        # Find all shapes (p:sp elements)
        for sp in root.iter():
            if sp.tag.endswith("}sp") or sp.tag == "sp":
                # Get shape properties using full namespace URIs
                nvSpPr = sp.find(f"{{{NS_P_URI}}}nvSpPr")
                if nvSpPr is None:
                    continue

                cNvPr = nvSpPr.find(f"{{{NS_P_URI}}}cNvPr")
                if cNvPr is None:
                    continue

                shape_id = int(cNvPr.get("id", 0))
                shape_name = cNvPr.get("name", "")

                # Get text from txBody
                txBody = sp.find(f"{{{NS_P_URI}}}txBody")
                if txBody is not None:
                    # Collect all text from runs
                    texts = []
                    # Iterate over p elements (paragraphs) in txBody
                    for p_elem in txBody:
                        if p_elem.tag.endswith("}p") or p_elem.tag == "p":
                            # Iterate over r elements (runs) in each paragraph
                            for r in p_elem:
                                if r.tag.endswith("}r") or r.tag == "r":
                                    for t in r:
                                        if t.tag.endswith("}t") or t.tag == "t":
                                            if t.text:
                                                texts.append(t.text)

                    if texts:
                        full_text = "".join(texts)
                        box = TextBox(slide_path, shape_id, shape_name, full_text)
                        text_boxes.append(box)

        return text_boxes

    def extract_text_boxes(self) -> TextBoxCollection:
        """Extract all text boxes from all slides.

        Returns:
            TextBoxCollection containing all found text boxes.
        """
        collection = TextBoxCollection()

        for slide_path in self.get_slide_files():
            boxes = self._extract_text_boxes_from_slide(slide_path)
            for box in boxes:
                collection.add(box)

        return collection

    def apply_translations(self, boxes: TextBoxCollection) -> None:
        """Apply translated text to the text boxes in the PPTX.

        Also handles font settings and autoFit based on configuration.

        Args:
            boxes: TextBoxCollection with updated text.
        """
        import xml.etree.ElementTree as ET

        # Track which slides have been modified
        modified_slides = set()

        for box in boxes:
            if box.text != box.original_text:
                modified_slides.add(box.slide_path)

        for slide_path in modified_slides:
            self._apply_translations_to_slide(slide_path, boxes)

    def _apply_translations_to_slide(self, slide_path: Path, boxes: TextBoxCollection) -> None:
        """Apply translations to a single slide.

        Args:
            slide_path: Path to the slide XML file.
            boxes: TextBoxCollection with translations.
        """
        # Read and parse the slide
        tree = ET.parse(slide_path)
        root = tree.getroot()

        # Get boxes for this slide
        slide_boxes = boxes.get_by_slide(slide_path)

        # Define namespace URIs
        NS_P_URI = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"

        for sp in root.iter():
            if sp.tag.endswith("}sp") or sp.tag == "sp":
                nvSpPr = sp.find(f"{{{NS_P_URI}}}nvSpPr")
                if nvSpPr is None:
                    continue

                cNvPr = nvSpPr.find(f"{{{NS_P_URI}}}cNvPr")
                if cNvPr is None:
                    continue

                shape_id = int(cNvPr.get("id", 0))

                # Find the matching box
                matching_box = None
                for box in slide_boxes:
                    if box.shape_id == shape_id:
                        matching_box = box
                        break

                if matching_box is None:
                    continue

                txBody = sp.find(f"{{{NS_P_URI}}}txBody")
                if txBody is None:
                    continue

                # Set western fonts for all text runs
                self._set_western_font(txBody)

                # Handle autoFit
                self._handle_autofit(txBody, matching_box)

                # Update text content
                self._update_text_content(txBody, matching_box.text)

        # Write the modified slide back
        tree.write(slide_path, encoding="UTF-8", xml_declaration=True)

    def _set_western_font(self, txBody) -> None:
        """Set western fonts (latin, cs, ea) to the default font.

        Args:
            txBody: The text body element.
        """
        default_font = self.config.default_font
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

        # Iterate over p elements (paragraphs), then r elements (runs)
        for p_elem in txBody:
            if p_elem.tag.endswith("}p") or p_elem.tag == "p":
                for r in p_elem:
                    if r.tag.endswith("}r") or r.tag == "r":
                        rPr = r.find(f"{{{NS_A}}}rPr")
                        if rPr is None:
                            rPr = ET.SubElement(r, f"{{{NS_A}}}rPr")

                        # Set latin font
                        latin = rPr.find(f"{{{NS_A}}}latin")
                        if latin is None:
                            latin = ET.SubElement(rPr, f"{{{NS_A}}}latin")
                        latin.set("typeface", default_font)

                        # Set cs (Complex Script) font
                        cs = rPr.find(f"{{{NS_A}}}cs")
                        if cs is None:
                            cs = ET.SubElement(rPr, f"{{{NS_A}}}cs")
                        cs.set("typeface", default_font)

                        # Set ea (East Asian) font
                        ea = rPr.find(f"{{{NS_A}}}ea")
                        if ea is None:
                            ea = ET.SubElement(rPr, f"{{{NS_A}}}ea")
                        ea.set("typeface", default_font)

    def _handle_autofit(self, txBody, box: TextBox) -> None:
        """Handle autoFit for text body based on text length changes.

        If the translated text is longer than the original, set fit="text"
        and adjust font size based on configuration.

        Args:
            txBody: The text body element.
            box: The TextBox with translation.
        """
        original_len = len(box.original_text) if box.original_text else 0
        new_len = len(box.text) if box.text else 0

        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

        # Find or create bodyPr
        bodyPr = txBody.find(f"{{{NS_A}}}bodyPr")
        if bodyPr is None:
            bodyPr = ET.Element(f"{{{NS_A}}}bodyPr")
            txBody.insert(0, bodyPr)

        # If new text is longer, apply autoFit
        if new_len > original_len and original_len > 0:
            # Set fit="text" to shrink text to fit
            bodyPr.set("fit", "text")

            # Calculate new font size based on ratio
            shrink_ratio = self.config.font_size_shrink_ratio
            font_size_min = self.config.font_size_min

            # Find rPr elements and adjust font size - iterate over p then r
            for p_elem in txBody:
                if p_elem.tag.endswith("}p") or p_elem.tag == "p":
                    for r in p_elem:
                        if r.tag.endswith("}r") or r.tag == "r":
                            rPr = r.find(f"{{{NS_A}}}rPr")
                            if rPr is not None:
                                # Get current font size if set
                                sz = rPr.get("sz")
                                if sz:
                                    current_size = int(sz)
                                    # Apply shrink ratio
                                    new_size = int(current_size * shrink_ratio)
                                    new_size = max(new_size, int(font_size_min * 100))  # font size in hundredths of a point
                                    rPr.set("sz", str(new_size))
        else:
            # Remove fit attribute if text got shorter or stayed same
            if "fit" in bodyPr.attrib:
                del bodyPr.attrib["fit"]

    def _update_text_content(self, txBody, new_text: str) -> None:
        """Update text content in a text body.

        Args:
            txBody: The text body element.
            new_text: The new text to set.
        """
        # Build list of text parts
        text_parts = []
        current_pos = 0

        # Collect all existing text runs - iterate over p then r
        runs = []
        for p_elem in txBody:
            if p_elem.tag.endswith("}p") or p_elem.tag == "p":
                for r in p_elem:
                    if r.tag.endswith("}r") or r.tag == "r":
                        runs.append(r)

        if not runs:
            return

        # Get text from each run and calculate positions
        for r in runs:
            for t in r:
                if t.tag.endswith("}t") or t.tag == "t":
                    if t.text:
                        text_parts.append((current_pos, current_pos + len(t.text), t))
                        current_pos += len(t.text)

        if not text_parts:
            return

        # If single text run, just update it
        if len(text_parts) == 1:
            _, _, t_elem = text_parts[0]
            t_elem.text = new_text
            return

        # Multiple runs: distribute text across runs
        # First run gets beginning, last run gets end, middle runs are cleared
        total_len = sum(len(t.text or "") for _, _, t in text_parts)

        if total_len == 0:
            text_parts[0][2].text = new_text
            return

        # Distribute new_text across runs
        remaining = new_text
        for i, (start, end, t_elem) in enumerate(text_parts):
            if i == 0:
                # First run: take from beginning
                part_len = end - start
                t_elem.text = remaining[:part_len] if remaining else ""
                remaining = remaining[part_len:]
            elif i == len(text_parts) - 1:
                # Last run: take remaining
                t_elem.text = remaining
            else:
                # Middle runs: clear text
                t_elem.text = ""
