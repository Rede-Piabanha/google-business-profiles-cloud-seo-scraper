# Security Policy

## Reporting a vulnerability

Please do not report security vulnerabilities through public GitHub issues.

Use GitHub Private Vulnerability Reporting when available, or contact Rede Piabanha through https://piabanha.net/.

Do not include credentials, API keys, personal data, exported lead data or other sensitive information in public reports.

## Credentials

This project expects Google API credentials to be supplied through environment variables. Never commit a populated `.env` file, API key, Programmable Search Engine CX intended to remain private, generated lead spreadsheets or local cache files.

If a credential is accidentally exposed, revoke or rotate it immediately in the relevant Google Cloud project and remove it from Git history before relying on deletion from the latest commit alone.
