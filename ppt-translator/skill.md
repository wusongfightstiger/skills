# PPT Translator

Translate Chinese PPT files to English using MiniMax API with terminology management.

## Usage

```
/ppt-translator <path-to-pptx> [--output-dir <dir>] [--glossary <path>]
```

## Options

- `<path-to-pptx>`: Path to the Chinese PPTX file to translate
- `--output-dir`, `-o`: Output directory (default: source-dir/Translated/)
- `--glossary`, `-g`: Path to custom glossary CSV

## Examples

```
/ppt-translator 中文/第1章.pptx
/ppt-translator 中文/第2章.pptx --output-dir /path/to/output
```

## How It Works

1. Unpacks the PPTX file and extracts text boxes
2. Each text box is translated independently via MiniMax API
3. Pre-replaces known terminology from the glossary
4. Sets Western fonts (Arial/Calibri) on translated text
5. Enables autoFit for character expansion handling
6. Repacks the translated content into a new PPTX
7. Outputs to `<source-dir>/Translated/<filename>_英文版.pptx`

## Terminology

The tool maintains a terminology glossary at:
`~/.claude/skills/ppt-translator/电路术语表.csv`

The glossary contains electrical engineering terminology (373+ terms).
New terms discovered during translation can be added to the glossary.

## Translation Rules

- Technical terms are pre-replaced using the glossary before API translation
- Formulas (E=mc², U=IR) are preserved unchanged
- Numbered lists (1-1, (a), (b), ① ② ③) are preserved
- Chinese person names are converted to Pinyin

## Error Handling

- Text box level rollback: if a text box fails, it reverts to original
- Other text boxes continue processing
- Failed text boxes are reported in the summary

## Environment Variables

- `MINIMAX_API_KEY`: Your MiniMax API key (required)

## Requirements

- Python 3.10+
- click
- requests
