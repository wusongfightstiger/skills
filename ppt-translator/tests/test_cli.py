"""Tests for CLI module."""

import pytest
from unittest.mock import patch, MagicMock
from click.testing import CliRunner

from ppt_translator.cli import cli, translate, status


class TestTranslateCommand:
    """Test cases for translate command."""

    def test_translate_command_without_api_key(self, test_pptx, monkeypatch):
        """Test that translate command fails without API key.

        When MINIMAX_API_KEY is not set, the command should exit with an error.
        """
        # Remove API key if set
        monkeypatch.delenv("MINIMAX_API_KEY", raising=False)

        runner = CliRunner()
        result = runner.invoke(translate, [str(test_pptx)])

        # Should fail due to missing API key
        assert result.exit_code != 0
        assert "MINIMAX_API_KEY" in result.output or "Error" in result.output

    def test_translate_command_with_mocked_api(self, test_pptx, test_glossary_csv, monkeypatch, tmp_path):
        """Test translate command with mocked API."""
        # Set dummy API key
        monkeypatch.setenv("MINIMAX_API_KEY", "test-api-key")

        # Mock the translator to avoid actual API calls
        with patch("ppt_translator.cli.MiniMaxTranslator") as mock_translator_class:
            mock_translator = MagicMock()
            mock_translator.translate.return_value = "translated text"
            mock_translator_class.return_value = mock_translator

            # Create output directory
            output_dir = tmp_path / "output"
            output_dir.mkdir()

            runner = CliRunner()
            result = runner.invoke(translate, [
                str(test_pptx),
                "-o", str(output_dir),
                "-g", str(test_glossary_csv)
            ])

            # Check command completed
            assert result.exit_code == 0 or "No text boxes found" in result.output


class TestStatusCommand:
    """Test cases for status command."""

    def test_status_command(self, test_glossary_csv):
        """Test status command displays glossary information."""
        runner = CliRunner()
        result = runner.invoke(status, ["-g", str(test_glossary_csv)])

        assert result.exit_code == 0
        # Should display glossary status
        assert "Terminology Glossary Status" in result.output or "Glossary" in result.output

    def test_status_command_with_nonexistent_glossary(self):
        """Test status command with non-existent glossary file."""
        runner = CliRunner()
        result = runner.invoke(status, ["-g", "/nonexistent/path/glossary.csv"])

        # Should handle missing file gracefully
        assert "not found" in result.output or result.exit_code != 0


class TestCliGroup:
    """Test cases for CLI group."""

    def test_cli_help(self):
        """Test that CLI help displays available commands."""
        runner = CliRunner()
        result = runner.invoke(cli, ["--help"])

        assert result.exit_code == 0
        assert "translate" in result.output
        assert "status" in result.output

    def test_cli_translate_help(self):
        """Test that translate command help is displayed."""
        runner = CliRunner()
        result = runner.invoke(translate, ["--help"])

        assert result.exit_code == 0
        assert "INPUT_PPTX" in result.output
        assert "--output-dir" in result.output or "-o" in result.output
        assert "--glossary" in result.output or "-g" in result.output

    def test_cli_status_help(self):
        """Test that status command help is displayed."""
        runner = CliRunner()
        result = runner.invoke(status, ["--help"])

        assert result.exit_code == 0
        assert "--glossary" in result.output or "-g" in result.output