"""CLI interface for PPT Translator."""

import click
import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Tuple, List, Dict

from .config import Config, get_api_key, get_glossary_path
from .terminology import TerminologyManager
from .translator import MiniMaxTranslator


class TextBox:
    """Represents a text box in a PPTX slide."""

    def __init__(self, slide_path: Path, xpath: str, original_text: str):
        """Initialize a TextBox.

        Args:
            slide_path: Path to the slide XML file.
            xpath: XPath to the text element within the slide.
            original_text: The original Chinese text.
        """
        self.slide_path = slide_path
        self.xpath = xpath
        self.original_text = original_text
        self.translated_text: Optional[str] = None
        self.status = "pending"  # pending, translated, failed
        self.backup_text: Optional[str] = None

    def mark_translated(self, translated_text: str) -> None:
        """Mark text box as successfully translated.

        Args:
            translated_text: The translated text.
        """
        self.translated_text = translated_text
        self.status = "translated"

    def mark_failed(self) -> None:
        """Mark text box translation as failed."""
        self.status = "failed"

    def rollback(self) -> None:
        """Rollback to original text if translation failed."""
        self.translated_text = None
        self.status = "pending"


class PPTXProcessor:
    """Processor for unpacking, extracting, and repacking PPTX files."""

    def __init__(self, pptx_path: Path):
        """Initialize PPTX processor.

        Args:
            pptx_path: Path to the PPTX file.
        """
        self.pptx_path = Path(pptx_path)
        self.temp_dir: Optional[Path] = None

    def unpack(self, dest_dir: Path) -> None:
        """Unpack PPTX to a directory.

        Args:
            dest_dir: Destination directory for unpacking.
        """
        dest_dir = Path(dest_dir)
        with zipfile.ZipFile(self.pptx_path, "r") as zf:
            zf.extractall(dest_dir)
        self.temp_dir = dest_dir

    def pack(self, output_path: Path) -> None:
        """Pack directory back to PPTX.

        Args:
            output_path: Output path for the PPTX file.
        """
        output_path = Path(output_path)
        if self.temp_dir is None:
            raise RuntimeError("PPTX not unpacked")

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in self.temp_dir.rglob("*"):
                if file_path.is_file():
                    arcname = file_path.relative_to(self.temp_dir)
                    zf.write(file_path, arcname)

    def extract_text_boxes(self, slides_dir: Path) -> List[TextBox]:
        """Extract text boxes from slide XML files.

        Args:
            slides_dir: Path to the slides directory.

        Returns:
            List of TextBox objects.
        """
        import xml.etree.ElementTree as ET

        text_boxes: List[TextBox] = []
        slides_dir = Path(slides_dir)

        # Namespace URIs
        NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
        NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

        for slide_path in sorted(slides_dir.glob("slide*.xml")):
            tree = ET.parse(slide_path)
            root = tree.getroot()

            # Find all sp (shape) elements in the slide
            for sp in root.iter(f"{NS_P}sp"):
                # Get the shape id from nvSpPr/cNvPr
                nvSpPr = sp.find(f"{NS_P}nvSpPr")
                if nvSpPr is None:
                    continue
                cNvPr = nvSpPr.find(f"{NS_P}cNvPr")
                if cNvPr is None:
                    continue
                sp_id = cNvPr.get("id", "unknown")

                # Find txBody within this shape
                txBody = sp.find(f"{NS_P}txBody")
                if txBody is None:
                    continue

                # Collect all text from paragraphs (p) and runs (r)
                full_text_parts = []
                for p_elem in txBody:
                    if p_elem.tag == f"{NS_A}p":
                        for r in p_elem:
                            if r.tag == f"{NS_A}r":
                                for t in r:
                                    if t.tag == f"{NS_A}t" and t.text:
                                        full_text_parts.append(t.text)

                if full_text_parts:
                    full_text = "".join(full_text_parts)
                    if full_text.strip():
                        xpath = f"{slide_path.name}//sp[@id='{sp_id}']/txBody"
                        text_box = TextBox(slide_path, xpath, full_text)
                        text_boxes.append(text_box)

        return text_boxes

    def apply_translation(self, text_box: TextBox, translated_text: str) -> None:
        """Apply translation to a text box in the slide XML.

        Args:
            text_box: The TextBox to update.
            translated_text: The translated text.
        """
        slide_path = self.temp_dir / "ppt" / "slides" / text_box.slide_path.name if self.temp_dir else text_box.slide_path

        # Register namespaces to preserve them
        namespaces = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }

        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)

        tree = ET.parse(slide_path)
        root = tree.getroot()

        # Find the shape by xpath stored in text_box
        # xpath format: "slideX.xml//sp[@id='Y']/txBody"
        shape_id = None
        for part in text_box.xpath.split("sp"):
            if "@id='" in part:
                start = part.find("@id='") + 5
                end = part.find("'", start)
                if end != -1:
                    shape_id = part[start:end]
                    break

        if not shape_id:
            # Fallback: try to find by original text
            for t_elem in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                if t_elem.text == text_box.original_text:
                    t_elem.text = translated_text
                    break
            tree.write(slide_path, encoding="UTF-8", xml_declaration=True)
            return

        # Find the specific shape (sp) element by shape id
        target_sp = None
        NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

        for sp in root.iter(f"{NS_P}sp"):
            nvSpPr = sp.find(f"{NS_P}nvSpPr")
            if nvSpPr is None:
                continue
            cNvPr = nvSpPr.find(f"{NS_P}cNvPr")
            if cNvPr is None:
                continue
            if cNvPr.get("id") == shape_id:
                target_sp = sp
                break

        if target_sp is None:
            return

        # Find txBody in this shape
        txBody = target_sp.find(f"{NS_P}txBody")
        if txBody is None:
            return

        # Set Western fonts (Arial) on all runs
        self._set_western_font(txBody)

        # Handle autoFit for character expansion
        self._handle_autofit(txBody, text_box.original_text, translated_text)

        # Collect all text runs (r elements) from all paragraphs
        NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        runs = []
        for p_elem in txBody:
            if p_elem.tag == f"{NS_A}p":
                for r in p_elem:
                    if r.tag == f"{NS_A}r":
                        runs.append(r)

        if not runs:
            return

        # Distribute translated text across runs
        # Strategy: put translated text in first run, clear others
        if len(runs) == 1:
            # Simple case: single run
            for t in runs[0]:
                if t.tag == f"{NS_A}t":
                    t.text = translated_text
                    break
        else:
            # Multi-run case: put text in first run, clear others
            # First run gets the full translated text
            first_run_has_text = False
            for t in runs[0]:
                if t.tag == f"{NS_A}t":
                    t.text = translated_text
                    first_run_has_text = True
                    break
            if not first_run_has_text:
                # Create t element if not exists
                t_elem = ET.SubElement(runs[0], f"{NS_A}t")
                t_elem.text = translated_text

            # Clear text in other runs
            for r in runs[1:]:
                for t in r:
                    if t.tag == f"{NS_A}t":
                        t.text = ""

        tree.write(slide_path, encoding="UTF-8", xml_declaration=True)

    def _set_western_font(self, txBody) -> None:
        """Set western fonts (latin, cs, ea) to Arial.

        Args:
            txBody: The text body element.
        """
        default_font = "Arial"
        NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

        def _set_font_on_rPr(rPr):
            """Helper to set fonts on an rPr element."""
            if rPr is None:
                return

            # Set latin font
            latin = rPr.find(f"{NS_A}latin")
            if latin is None:
                latin = ET.SubElement(rPr, f"{NS_A}latin")
            latin.set("typeface", default_font)

            # Set cs (Complex Script) font
            cs = rPr.find(f"{NS_A}cs")
            if cs is None:
                cs = ET.SubElement(rPr, f"{NS_A}cs")
            cs.set("typeface", default_font)

            # Set ea (East Asian) font
            ea = rPr.find(f"{NS_A}ea")
            if ea is None:
                ea = ET.SubElement(rPr, f"{NS_A}ea")
            ea.set("typeface", default_font)

        # Iterate over p elements (paragraphs)
        for p_elem in txBody:
            if p_elem.tag == f"{NS_A}p":
                # Process all run elements (r)
                for r in p_elem:
                    if r.tag == f"{NS_A}r":
                        rPr = r.find(f"{NS_A}rPr")
                        if rPr is None:
                            rPr = ET.SubElement(r, f"{NS_A}rPr")
                        _set_font_on_rPr(rPr)

                # Also process endParaRPr (end paragraph run properties)
                endParaRPr = p_elem.find(f"{NS_A}endParaRPr")
                if endParaRPr is not None:
                    _set_font_on_rPr(endParaRPr)

    def _handle_autofit(self, txBody, original_text: str, new_text: str) -> None:
        """Handle autoFit for text body based on text length changes.

        If the translated text is longer than the original, set fit="text"
        and adjust font size based on configuration.

        Args:
            txBody: The text body element.
            original_text: The original text before translation.
            new_text: The new translated text.
        """
        original_len = len(original_text) if original_text else 0
        new_len = len(new_text) if new_text else 0

        NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

        # Find or create bodyPr
        bodyPr = txBody.find(f"{NS_A}bodyPr")
        if bodyPr is None:
            bodyPr = ET.Element(f"{NS_A}bodyPr")
            txBody.insert(0, bodyPr)

        # Font size settings
        shrink_ratio = 0.5  # 50% shrink
        font_size_min = 8   # 8pt minimum

        # If new text is longer, apply autoFit
        if new_len > original_len and original_len > 0:
            # Set fit="text" to shrink text to fit
            bodyPr.set("fit", "text")

            # Find rPr elements and adjust font size
            for p_elem in txBody:
                if p_elem.tag == f"{NS_A}p":
                    for r in p_elem:
                        if r.tag == f"{NS_A}r":
                            rPr = r.find(f"{NS_A}rPr")
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


@click.group()
def cli():
    """PPT Translator CLI - Translate Chinese PPT to English."""
    pass


@cli.command()
@click.argument("input_pptx", type=click.Path(exists=True))
@click.option("--output-dir", "-o", type=click.Path(), help="Output directory for translated PPTX")
@click.option("--glossary", "-g", type=click.Path(exists=True), help="Path to glossary CSV file")
def translate(input_pptx: str, output_dir: Optional[str], glossary: Optional[str]):
    """Translate a PPTX file from Chinese to English.

    INPUT_PPTX: Path to the input PPTX file.

    The translation process:
    1. Unpack the PPTX file
    2. Extract text boxes from slides
    3. Translate each text box using MiniMax API
    4. Apply translations and repack the PPTX
    """
    input_path = Path(input_pptx)

    # Determine output path
    if output_dir:
        output_dir_path = Path(output_dir)
        # If output_dir is a .pptx file path, use it directly as output
        if output_dir_path.suffix.lower() == '.pptx':
            output_path = output_dir_path
        else:
            output_path = output_dir_path / input_path.name
    else:
        output_path = input_path.with_stem(f"{input_path.stem}_translated")

    # Ensure parent directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Initialize terminology manager
    glossary_path = Path(glossary) if glossary else get_glossary_path()
    if not glossary_path.exists():
        click.echo(f"Warning: Glossary file not found at {glossary_path}", err=True)

    term_manager = TerminologyManager(glossary_path)

    # Initialize translator
    try:
        api_key = get_api_key()
    except SystemExit:
        click.echo("Error: MINIMAX_API_KEY environment variable is not set.", err=True)
        raise SystemExit(1)

    translator = MiniMaxTranslator(api_key)

    # Create temp directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        unpack_dir = temp_path / "unpacked"

        # Step 1: Unpack PPTX
        click.echo(f"Unpacking {input_path}...")
        processor = PPTXProcessor(input_path)
        processor.unpack(unpack_dir)

        # Step 2: Extract text boxes
        slides_dir = unpack_dir / "ppt" / "slides"
        click.echo("Extracting text boxes...")
        text_boxes = processor.extract_text_boxes(slides_dir)
        click.echo(f"Found {len(text_boxes)} text boxes")

        if not text_boxes:
            click.echo("No text boxes found to translate.")
            processor.pack(output_path)
            click.echo(f"Output saved to: {output_path}")
            return

        # Step 3: Translate each text box
        success_count = 0
        failed_count = 0

        click.echo("Translating...")

        for i, text_box in enumerate(text_boxes, 1):
            # Pre-replace terminology
            preprocessed = term_manager.pre_replace(text_box.original_text)

            try:
                # Translate using API
                translated = translator.translate(preprocessed, source_lang="zh", target_lang="en")
                text_box.mark_translated(translated)
                success_count += 1

                # Apply translation to the slide
                processor.apply_translation(text_box, translated)

                if i % 10 == 0:
                    click.echo(f"Progress: {i}/{len(text_boxes)}")

            except Exception as e:
                text_box.mark_failed()
                failed_count += 1
                click.echo(f"Failed to translate '{text_box.original_text[:30]}...': {e}")
                # Rollback - the translation is not applied on failure

        # Step 4: Pack output
        click.echo("Packing output PPTX...")
        processor.pack(output_path)

    # Step 5: Display summary
    click.echo("\n" + "=" * 50)
    click.echo("Translation Summary")
    click.echo("=" * 50)
    click.echo(f"Total text boxes: {len(text_boxes)}")
    click.echo(f"Successfully translated: {success_count}")
    click.echo(f"Failed: {failed_count}")
    click.echo(f"Output file: {output_path}")

    # Step 6: Display new discovered terms
    new_terms = term_manager.get_new_terms_summary()
    if new_terms:
        click.echo("\n" + "-" * 50)
        click.echo("New Discovered Terms")
        click.echo("-" * 50)
        for term in new_terms:
            click.echo(f"  {term['中文术语']} -> {term['英文翻译']}")


@cli.command()
@click.option("--glossary", "-g", type=click.Path(exists=True), help="Path to glossary CSV file")
def status(glossary: Optional[str]):
    """Display terminology glossary status.

    Shows information about the glossary including total terms and domains.
    """
    glossary_path = Path(glossary) if glossary else get_glossary_path()

    if not glossary_path.exists():
        click.echo(f"Glossary file not found at {glossary_path}", err=True)
        click.echo("Please specify a valid glossary file with --glossary/-g option.")
        return

    term_manager = TerminologyManager(glossary_path)

    # Count terms by domain
    domains: Dict[str, int] = {}
    for chinese, domain in term_manager._domains.items():
        domains[domain] = domains.get(domain, 0) + 1

    click.echo("=" * 50)
    click.echo("Terminology Glossary Status")
    click.echo("=" * 50)
    click.echo(f"Glossary file: {glossary_path}")
    click.echo(f"Total confirmed terms: {len(term_manager._terms)}")
    click.echo(f"New undiscovered terms: {len(term_manager.get_new_terms_summary())}")

    if domains:
        click.echo("\nTerms by domain:")
        for domain, count in sorted(domains.items()):
            click.echo(f"  {domain}: {count}")


if __name__ == "__main__":
    cli()