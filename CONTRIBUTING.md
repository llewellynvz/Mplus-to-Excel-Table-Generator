# Contributing

## Scope
This repository focuses on parsing Mplus measurement model output files and producing APA-style Excel tables.

## Guidelines
- Keep analytic functionality stable.
- Prefer additive changes (logging, diagnostics, test coverage, packaging).
- If you change parsing behavior, document it clearly and add tests.

## Development setup
```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate # Windows
pip install -r requirements.txt
```
