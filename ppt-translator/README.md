# PPT Translator

Translate Chinese PPT files to English with terminology management.

## Quick Start

1. Set your MiniMax API key:
```bash
export MINIMAX_API_KEY='your-api-key'
```

2. Translate a PPT file:
```bash
python -m src.ppt_translator.cli translate 中文/第1章.pptx
```

3. Or use the Claude Skill:
```
/ppt-translator 中文/第1章.pptx
```

## Project Structure

```
ppt-translator/
├── src/ppt_translator/
│   ├── cli.py           # CLI interface
│   ├── config.py        # Configuration
│   ├── terminology.py    # Terminology management
│   ├── text_box.py      # Text box model
│   ├── translator.py     # MiniMax API
│   └── ppt_handler.py    # PPTX operations
├── tests/               # Tests
├── 电路术语表.csv       # Terminology glossary (373+ terms)
└── docs/               # Documentation
```

## Terminology Glossary

The glossary file `电路术语表.csv` contains pre-loaded electrical engineering terminology.

Format:
```csv
中文术语,英文翻译,领域,添加日期,是否已确认
电阻,resistor,电路,2026-03-21,是
```

## Development

```bash
# Install dependencies
pip install click requests pytest pytest-mock

# Run tests
pytest tests/ -v

# Format code
black src/
```

## Translation Process

1. **Unpack**: PPTX is extracted to temporary directory
2. **Extract Text Boxes**: Each shape with text is extracted as a TextBox
3. **Pre-replace**: Known terms from glossary are replaced before API call
4. **Translate**: Each text box is translated via MiniMax API
5. **Apply**: Translated text is inserted, fonts set to Arial/Calibri
6. **Pack**: Result is packaged as new PPTX

## Error Handling

- Text box level: if one text box fails, only that box is rolled back
- Batch level: other text boxes continue processing
- Summary reports succeeded/failed counts

## License

MIT
