"""Tests for config module."""

import pytest
from unittest.mock import patch

from ppt_translator.config import Config, get_api_key, get_glossary_path


class TestConfig:
    """Test cases for Config class."""

    def test_config_loads_api_key(self, monkeypatch):
        """Test that Config loads api_key from MINIMAX_API_KEY environment variable."""
        test_api_key = "test-api-key-12345"
        monkeypatch.setenv("MINIMAX_API_KEY", test_api_key)
        config = Config()
        assert config.api_key == test_api_key

    def test_config_glossary_path(self, monkeypatch):
        """Test that Config().glossary_path contains the default glossary filename."""
        monkeypatch.setenv("MINIMAX_API_KEY", "dummy-key")
        config = Config()
        assert "电路术语表.csv" in str(config.glossary_path)


class TestGetApiKey:
    """Test cases for get_api_key function."""

    def test_get_api_key_raises_without_key(self, monkeypatch):
        """Test that get_api_key raises SystemExit when MINIMAX_API_KEY is not set."""
        monkeypatch.delenv("MINIMAX_API_KEY", raising=False)
        with pytest.raises(SystemExit) as exc_info:
            get_api_key()
        assert exc_info.value.code == 1


class TestGetGlossaryPath:
    """Test cases for get_glossary_path function."""

    def test_get_glossary_path_returns_path_object(self):
        """Test that get_glossary_path returns a Path object."""
        path = get_glossary_path()
        assert isinstance(path, type(get_glossary_path()))

    def test_get_glossary_path_contains_glossary_filename(self):
        """Test that glossary path contains the expected filename."""
        path = get_glossary_path()
        assert "电路术语表.csv" in str(path)
