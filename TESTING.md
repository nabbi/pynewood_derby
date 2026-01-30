# Testing

## Setup

Install dependencies:

```bash
pip install -r requirements.txt -r requirements_dev.txt
```

## Run tests

```bash
python -m pytest
```

## Run with coverage

```bash
python -m pytest --cov --cov-report=term-missing
```

## CI

Tests run automatically on push/PR to `main` via GitHub Actions across Python 3.9–3.12. Coverage is reported on the 3.12 run.
