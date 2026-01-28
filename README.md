# xlsagent

A collection of Claude Code skills for working with Excel workbooks and related data.

## Skills

### excel-analyzer

Comprehensive Excel workbook analyzer that extracts and documents:
- Sheet metadata (visibility, dimensions, features)
- Formulas with classification (LAMBDA, dynamic array, lookup, etc.)
- Named ranges and LAMBDA function definitions
- VBA macros
- Power Query M code
- Conditional formatting rules
- Data validations
- Pivot tables and charts
- Structured tables
- And much more...

Generates both:
- **Rich HTML report** for human review (searchable, navigable)
- **Agent-optimized Markdown files** for AI consumption

See [skills/excel-analyzer/SKILL.md](skills/excel-analyzer/SKILL.md) for usage details.

## Installation

Each skill has its own Python environment. To set up:

```bash
cd skills/excel-analyzer
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
playwright install chromium
```

## Running Tests

```bash
cd skills/excel-analyzer
source .venv/bin/activate
pytest tests/ -v
```

## Requirements

- Python 3.11+
- For SharePoint downloads: Playwright browsers

## License

MIT
