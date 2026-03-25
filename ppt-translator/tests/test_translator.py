"""Tests for translator.py."""

from unittest.mock import patch, MagicMock
import pytest

from ppt_translator.translator import MiniMaxTranslator


class TestMiniMaxTranslator:
    """Test cases for MiniMaxTranslator."""

    def _setup_mock_config(self, mock_config_cls):
        """Setup mock Config class."""
        mock_config = MagicMock()
        mock_config.minimax_api_host = "api.minimax.chat"
        mock_config.minimax_model = "MiniMax-M2.7"
        mock_config.request_timeout = 120
        mock_config_cls.return_value = mock_config

    @patch("ppt_translator.translator.Config")
    @patch("ppt_translator.translator.requests.post")
    def test_translate_text(self, mock_post, mock_config_cls):
        """Test translate text: mock API returns 'Hello', verify translation of '你好'."""
        self._setup_mock_config(mock_config_cls)

        # Setup mock response - Anthropic API format
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "content": [
                {"type": "text", "text": "Hello"}
            ]
        }
        mock_post.return_value = mock_response

        # Create translator and translate
        translator = MiniMaxTranslator(api_key="test-api-key")
        result = translator.translate("你好", source_lang="zh", target_lang="en")

        # Verify result
        assert result == "Hello"

    @patch("ppt_translator.translator.Config")
    @patch("ppt_translator.translator.requests.post")
    def test_translate_calls_api_with_correct_params(self, mock_post, mock_config_cls):
        """Test that API is called with correct parameters."""
        self._setup_mock_config(mock_config_cls)

        # Setup mock response - Anthropic API format
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "content": [
                {"type": "text", "text": "Test translation"}
            ]
        }
        mock_post.return_value = mock_response

        # Create translator and call translate
        translator = MiniMaxTranslator(api_key="my-secret-key")
        translator.translate("测试文本", source_lang="zh", target_lang="en")

        # Verify the API was called with correct parameters
        mock_post.assert_called_once()
        call_kwargs = mock_post.call_args.kwargs

        # Check headers
        assert "Authorization" in call_kwargs["headers"]
        assert call_kwargs["headers"]["Authorization"] == "Bearer my-secret-key"
        assert call_kwargs["headers"]["anthropic-version"] == "2023-06-01"

        # Check body
        assert "json" in call_kwargs
        body = call_kwargs["json"]
        assert body["model"] == "MiniMax-M2.7"
        assert body["max_tokens"] == 1024
        assert len(body["messages"]) == 1
        assert body["messages"][0]["role"] == "user"
        assert "Translate the following text from zh to en" in body["messages"][0]["content"]
        assert "测试文本" in body["messages"][0]["content"]

    @patch("ppt_translator.translator.Config")
    def test_translate_empty_text(self, mock_config_cls):
        """Test that empty text returns original text."""
        self._setup_mock_config(mock_config_cls)

        translator = MiniMaxTranslator(api_key="test-api-key")

        # Test empty string
        result = translator.translate("", source_lang="zh", target_lang="en")
        assert result == ""

        # Test whitespace only
        result = translator.translate("   ", source_lang="zh", target_lang="en")
        assert result == "   "

    @patch("ppt_translator.translator.Config")
    @patch("ppt_translator.translator.requests.post")
    def test_translate_timeout(self, mock_post, mock_config_cls):
        """Test timeout handling."""
        self._setup_mock_config(mock_config_cls)

        import requests

        # Setup mock to raise timeout exception
        mock_post.side_effect = requests.exceptions.Timeout("Connection timed out")

        translator = MiniMaxTranslator(api_key="test-api-key")

        # Verify TimeoutError is raised
        with pytest.raises(TimeoutError) as exc_info:
            translator.translate("测试", source_lang="zh", target_lang="en")

        assert "timed out" in str(exc_info.value)

    @patch("ppt_translator.translator.Config")
    @patch("ppt_translator.translator.requests.post")
    def test_translate_api_error(self, mock_post, mock_config_cls):
        """Test API error handling."""
        self._setup_mock_config(mock_config_cls)

        import requests

        # Setup mock to raise connection error
        mock_post.side_effect = requests.exceptions.ConnectionError("Connection refused")

        translator = MiniMaxTranslator(api_key="test-api-key")

        # Verify RuntimeError is raised
        with pytest.raises(RuntimeError) as exc_info:
            translator.translate("测试", source_lang="zh", target_lang="en")

        assert "failed" in str(exc_info.value)

    @patch("ppt_translator.translator.Config")
    def test_build_prompt(self, mock_config_cls):
        """Test prompt building."""
        self._setup_mock_config(mock_config_cls)

        translator = MiniMaxTranslator(api_key="test-api-key")
        prompt = translator._build_prompt("Hello", "en", "zh")

        assert "Translate the following text from en to zh" in prompt
        assert "Hello" in prompt

    @patch("ppt_translator.translator.Config")
    def test_clean_translation(self, mock_config_cls):
        """Test translation cleaning."""
        self._setup_mock_config(mock_config_cls)

        translator = MiniMaxTranslator(api_key="test-api-key")

        # Test basic cleaning
        assert translator._clean_translation("  Hello  ") == "Hello"

        # Test quoted translation
        assert translator._clean_translation('"Hello"') == "Hello"
        assert translator._clean_translation("'Hello'") == "Hello"

        # Test empty
        assert translator._clean_translation("") == ""
