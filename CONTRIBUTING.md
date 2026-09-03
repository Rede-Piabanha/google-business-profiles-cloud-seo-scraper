# Contributing

Contributions are welcome.

## Development setup

1. Create and activate a Python virtual environment.
2. Install dependencies with:

```sh
python -m pip install -r requirements.txt
```

3. Copy `.env.example` to `.env` only when local API-backed testing is required. Never commit populated credentials.
4. Before opening a pull request, run:

```sh
python -m py_compile google-business-profiles-cloud-seo-scraper.py
python -m unittest discover -s tests -v
```

## Pull requests

Keep pull requests focused and explain what changed, why it changed and how it was tested.

Do not include API keys, credentials, generated spreadsheets, cached API responses, personal data, lead lists or other sensitive/exported data.

Changes that increase Google API usage, alter scoring logic, change data collection or add external services should be clearly documented in the pull request.

By contributing, you agree that your contribution may be distributed under the MIT License.
