"""Abstract base class for translation engines."""

from abc import ABC, abstractmethod


class TranslationEngine(ABC):
    """Base class for all translation engines."""

    @abstractmethod
    async def translate_slide(self, slide_data: dict, glossary: list[dict]) -> dict:
        """Translate a single slide's text elements.

        Args:
            slide_data: Dict with slide_number and elements (as extracted by pptx_engine).
            glossary: List of glossary terms [{"zh": ..., "en": ..., "domain": ...}].

        Returns:
            Dict with same structure, text fields replaced with translations.
        """
        pass

    @abstractmethod
    def name(self) -> str:
        """Return human-readable engine name."""
        pass
