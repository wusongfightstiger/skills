"""MiniMax translation engine using Anthropic-compatible API."""

import json
import os
import re

import httpx

from .base import TranslationEngine
from ..prompt import SYSTEM_PROMPT, build_slide_prompt
from ..glossary import extract_relevant_terms


class MiniMaxEngine(TranslationEngine):
    """Translation engine using MiniMax API."""

    def __init__(self, model: str = "MiniMax-M2.7", api_key: str | None = None):
        self.model = model
        self.api_key = api_key or os.environ.get("MINIMAX_API_KEY", "")
        self.api_url = "https://api.minimaxi.com/anthropic/v1/messages"
        self.timeout = 120

    def name(self) -> str:
        return f"minimax ({self.model})"

    async def translate_slide(self, slide_data: dict, glossary: list[dict]) -> dict:
        # Build prompt
        slide_text = " ".join(
            r["text"] for e in slide_data["elements"] for r in e["runs"]
        )
        matched_terms = extract_relevant_terms(slide_text, glossary) if glossary else []
        user_prompt = build_slide_prompt(slide_data, matched_terms)

        # Call API
        headers = {
            "Content-Type": "application/json",
            "x-api-key": self.api_key,
            "anthropic-version": "2024-05-01",
        }
        payload = {
            "model": self.model,
            "max_tokens": 8192,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": user_prompt}],
        }

        async with httpx.AsyncClient(timeout=self.timeout) as client:
            resp = await client.post(self.api_url, headers=headers, json=payload)
            resp.raise_for_status()

        result = resp.json()
        text_content = _extract_text_from_response(result)
        return _parse_translation_json(text_content, slide_data)


class ClaudeAPIEngine(TranslationEngine):
    """Translation engine using Anthropic Messages API."""

    def __init__(self, model: str = "claude-sonnet-4-6-20250514", api_key: str | None = None):
        self.model = model
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        self.api_url = "https://api.anthropic.com/v1/messages"
        self.timeout = 120

    def name(self) -> str:
        return f"claude-api ({self.model})"

    async def translate_slide(self, slide_data: dict, glossary: list[dict]) -> dict:
        slide_text = " ".join(
            r["text"] for e in slide_data["elements"] for r in e["runs"]
        )
        matched_terms = extract_relevant_terms(slide_text, glossary) if glossary else []
        user_prompt = build_slide_prompt(slide_data, matched_terms)

        headers = {
            "Content-Type": "application/json",
            "x-api-key": self.api_key,
            "anthropic-version": "2024-05-01",
        }
        payload = {
            "model": self.model,
            "max_tokens": 8192,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": user_prompt}],
        }

        async with httpx.AsyncClient(timeout=self.timeout) as client:
            resp = await client.post(self.api_url, headers=headers, json=payload)
            resp.raise_for_status()

        result = resp.json()
        text_content = _extract_text_from_response(result)
        return _parse_translation_json(text_content, slide_data)


# ── Shared helpers ─────────────────────────────────────────────────

def _extract_text_from_response(response: dict) -> str:
    """Extract text content from Anthropic-format API response."""
    for block in response.get("content", []):
        if block.get("type") == "text":
            return block.get("text", "")
    return ""


def _parse_translation_json(text: str, original_slide: dict) -> dict:
    """Parse LLM response text as JSON, with fallback cleaning."""
    text = text.strip()

    # Remove markdown code fence if present
    if text.startswith("```"):
        lines = text.split("\n")
        # Remove first line (```json) and last line (```)
        lines = [l for l in lines if not l.strip().startswith("```")]
        text = "\n".join(lines)

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        # Try to extract JSON object from the text
        match = re.search(r'\{[\s\S]*\}', text)
        if match:
            try:
                parsed = json.loads(match.group())
            except json.JSONDecodeError:
                # Return original as fallback
                return original_slide
        else:
            return original_slide

    # Validate structure
    if not _validate_translation(original_slide, parsed):
        return original_slide

    return parsed


def _validate_translation(original: dict, translated: dict) -> bool:
    """Validate that translated JSON matches original structure."""
    orig_elements = original.get("elements", [])
    trans_elements = translated.get("elements", [])

    if len(trans_elements) != len(orig_elements):
        return False

    for orig_elem, trans_elem in zip(orig_elements, trans_elements):
        if orig_elem.get("id") != trans_elem.get("id"):
            return False
        orig_runs = orig_elem.get("runs", [])
        trans_runs = trans_elem.get("runs", [])
        if len(trans_runs) != len(orig_runs):
            return False
        for tr in trans_runs:
            if not tr.get("text"):
                return False

    return True
