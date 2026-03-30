"""Text box model for granular tracking of text translations."""

from dataclasses import dataclass, field
from typing import Optional, Iterator


@dataclass
class TextBox:
    """Represents a single text box/shape in a PPT slide.

    Attributes:
        shape_id: Unique identifier for the shape
        shape_name: Name of the shape from the PPT
        original_text: The original text before translation
        xpath: XPath to locate this text box in the XML
        translated_text: The translated text (None if not yet translated)
        failed: Whether the translation failed
        error_message: Error message if translation failed
    """

    shape_id: int
    shape_name: str
    original_text: str
    xpath: str
    translated_text: Optional[str] = None
    failed: bool = False
    error_message: Optional[str] = None

    def is_translated(self) -> bool:
        """Check if the text has been successfully translated.

        Returns:
            True if translated_text is set and not failed
        """
        return self.translated_text is not None and not self.failed

    def mark_translated(self, translated_text: str) -> None:
        """Mark the text box as successfully translated.

        Args:
            translated_text: The translated text
        """
        self.translated_text = translated_text
        self.failed = False
        self.error_message = None

    def mark_failed(self, error_message: str) -> None:
        """Mark the text box as failed translation.

        Args:
            error_message: Description of the failure
        """
        self.failed = True
        self.error_message = error_message
        self.translated_text = None

    def rollback(self) -> None:
        """Rollback the translation, reverting to original text."""
        self.translated_text = None
        self.failed = False
        self.error_message = None

    def get_final_text(self) -> str:
        """Get the final text (translated or original).

        Returns:
            Translated text if available, otherwise original text
        """
        if self.is_translated():
            return self.translated_text
        return self.original_text


class TextBoxCollection:
    """Collection of TextBox objects for tracking all text boxes in a presentation."""

    def __init__(self):
        """Initialize an empty collection."""
        self._boxes: dict[int, TextBox] = {}

    def add(self, text_box: TextBox) -> None:
        """Add a text box to the collection.

        Args:
            text_box: TextBox to add
        """
        self._boxes[text_box.shape_id] = text_box

    def get_by_id(self, shape_id: int) -> Optional[TextBox]:
        """Get a text box by its shape_id.

        Args:
            shape_id: The shape identifier

        Returns:
            TextBox if found, None otherwise
        """
        return self._boxes.get(shape_id)

    def __len__(self) -> int:
        """Return the number of text boxes in the collection."""
        return len(self._boxes)

    def __iter__(self) -> Iterator[TextBox]:
        """Iterate over all text boxes in the collection."""
        return iter(self._boxes.values())

    def get_failed(self) -> list[TextBox]:
        """Get all text boxes that failed translation.

        Returns:
            List of failed TextBox objects
        """
        return [box for box in self._boxes.values() if box.failed]

    def get_successful(self) -> list[TextBox]:
        """Get all text boxes that were successfully translated.

        Returns:
            List of successfully translated TextBox objects
        """
        return [box for box in self._boxes.values() if box.is_translated()]

    def summary(self) -> dict[str, int]:
        """Get a summary of the collection status.

        Returns:
            Dictionary with counts: total, translated, failed, pending
        """
        total = len(self._boxes)
        translated = len(self.get_successful())
        failed = len(self.get_failed())
        pending = total - translated - failed

        return {
            "total": total,
            "translated": translated,
            "failed": failed,
            "pending": pending,
        }