"""MiniMax API translator module for PPT Translator."""

import requests
from typing import List, Optional

from .config import Config


class MiniMaxTranslator:
    """Translator using MiniMax API."""

    def __init__(self, api_key: str):
        """Initialize the translator.

        Args:
            api_key: MiniMax API key.
        """
        self.api_key = api_key
        self.config = Config()
        # Use Anthropic-compatible endpoint
        self.url = "https://api.minimaxi.com/anthropic/v1/messages"
        self.model = "MiniMax-M2.7"
        self.timeout = self.config.request_timeout

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Translate text from source language to target language.

        Args:
            text: Text to translate.
            source_lang: Source language code (e.g., "zh" for Chinese).
            target_lang: Target language code (e.g., "en" for English).

        Returns:
            Translated text.

        Raises:
            TimeoutError: When the request times out.
            RuntimeError: When the API returns an error.
        """
        if not text or not text.strip():
            return text

        prompt = self._build_prompt(text, source_lang, target_lang)

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01",
        }

        body = {
            "model": self.model,
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "max_tokens": 1024,
        }

        try:
            response = requests.post(
                self.url,
                headers=headers,
                json=body,
                timeout=self.timeout,
            )
            response.raise_for_status()
        except requests.exceptions.Timeout:
            raise TimeoutError(f"Request to MiniMax API timed out after {self.timeout}s")
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"MiniMax API request failed: {e}")

        result = response.json()
        # Anthropic API format: {"content": [{"type": "text", "text": "..."}]}
        content = result.get("content", [])
        if content and isinstance(content, list):
            # Find text type content, skip thinking
            for item in content:
                if item.get("type") == "text":
                    translated_text = item.get("text", "")
                    break
            else:
                translated_text = ""
        else:
            translated_text = ""

        return self._clean_translation(translated_text)

    def translate_batch(self, texts: List[str], source_lang: str, target_lang: str) -> List[str]:
        """Translate multiple texts.

        Args:
            texts: List of texts to translate.
            source_lang: Source language code.
            target_lang: Target language code.

        Returns:
            List of translated texts.
        """
        return [self.translate(text, source_lang, target_lang) for text in texts]

    def _build_prompt(self, text: str, source_lang: str, target_lang: str) -> str:
        """Build the translation prompt.

        Args:
            text: Text to translate.
            source_lang: Source language code.
            target_lang: Target language code.

        Returns:
            Formatted prompt string.
        """
        return f"Translate the following text from {source_lang} to {target_lang}. Only return the translated text, no explanations.\n\nText: {text}"

    def _clean_translation(self, translation: str) -> str:
        """Clean the translation result.

        Args:
            translation: Raw translation from API.

        Returns:
            Cleaned translation text.
        """
        if not translation:
            return ""

        # Remove leading/trailing whitespace
        translation = translation.strip()

        # Remove quotes if the entire translation is wrapped in them
        if translation.startswith('"') and translation.endswith('"'):
            translation = translation[1:-1].strip()
        elif translation.startswith("'") and translation.endswith("'"):
            translation = translation[1:-1].strip()

        return translation
