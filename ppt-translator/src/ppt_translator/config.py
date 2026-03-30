"""Configuration management for PPT Translator."""

import os
import sys
from pathlib import Path


# Project root directory (parent of src/ppt_translator)
PROJECT_ROOT = Path(__file__).parent.parent.parent

# Glossary directory and default glossary file
GLOSSARY_DIR = PROJECT_ROOT
DEFAULT_GLOSSARY = "电路术语表.csv"


class Config:
    """Configuration class for PPT Translator."""

    # MiniMax API settings
    minimax_api_host = "api.minimax.chat"
    minimax_model = "MiniMax-Text-01"

    # Request settings
    request_timeout = 120  # seconds

    # Font settings
    default_font = "Arial"
    font_size_min = 8  # minimum font size in points
    font_size_shrink_ratio = 0.5  # maximum shrink ratio (50%)

    def __init__(self):
        """Initialize configuration from environment variables."""
        self.api_key = get_api_key()
        self.glossary_path = get_glossary_path()


def get_api_key() -> str:
    """Get API key from environment variable MINIMAX_API_KEY.

    Prints error message and exits with code 1 if key is not found.
    """
    api_key = os.environ.get("MINIMAX_API_KEY")
    if not api_key:
        print("Error: MINIMAX_API_KEY environment variable is not set.", file=sys.stderr)
        raise SystemExit(1)
    return api_key


def get_glossary_path() -> Path:
    """Get the path to the default glossary CSV file.

    Returns:
        Path to GLOSSARY_DIR / DEFAULT_GLOSSARY
    """
    return GLOSSARY_DIR / DEFAULT_GLOSSARY
