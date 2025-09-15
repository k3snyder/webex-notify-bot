# Webex Notify Bot

Bulk‑send 1:1 Webex messages with an Adaptive Card using a bot token. Reads recipient emails from a CSV, personalizes a card template, sends in batches with retries, and appends results to a delivery log.

## Features
- Bulk 1:1 messaging with Adaptive Cards
- CSV input (`email` header or first column)
- Placeholder substitution in `main.json` (account, opportunity, amount, due, cta_url)
- Batching with inter‑batch delay to respect rate limits
- Per‑recipient retries with delay on failure
- Append‑only delivery log CSV with status and HTTP code
- Dry‑run mode to preview without API calls

## Requirements
- Python 3.10+
- `requests` (installed via `requirements.txt`)
- Webex Bot access token with permission to create messages

## Quick Start
```bash
# Create and activate a virtual environment
python3 -m venv .venv && source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Provide your Webex bot token (recommended via env var)
export WEBEX_BOT_TOKEN="<bot_token>"

# Optional: review defaults in settings.json
cat settings.json

# Dry run (no API calls)
python3 send_webex_notifications.py --csv recipients.csv --dry-run

# Use settings.json values (no flags)
python3 send_webex_notifications.py

# Send using sample fields
python3 send_webex_notifications.py \
  --csv recipients.csv \
  --account "ACME" \
  --opportunity "Q4 Expansion" \
  --amount "$50,000" \
  --due 2025-10-01 \
  --cta-url https://example.crm.com/opps/ACME-Q4
```

## Configuration
The script resolves configuration with this precedence:
1) CLI flag → 2) `settings.json` → 3) built‑in defaults

Supported keys in `settings.json` (snake_case):
```json
{
  "csv": "recipients.csv",
  "log_file": "send_log.csv",
  "batch_size": 10,
  "batch_delay": 5.0,
  "retry_count": 3,
  "retry_delay": 5.0,
  "account": "ACME",
  "opportunity": "Q4 Expansion",
  "amount": "$50,000",
  "due": "2025-10-01",
  "cta_url": "https://example.crm.com/opps/ACME-Q4",
  "card_json": "main.json"
}
```

Provide the bot token by one of:
- `--token <token>`
- Environment: `WEBEX_BOT_TOKEN`
- (Fallback) `settings.json` key `token` (not recommended for security)

## CSV Format
- Header `email` or first column of each row is treated as the address.
- Example `recipients.csv`:
```
email
user1@example.com
user2@example.com
```
- Rows with empty/invalid emails are skipped. Duplicates are removed while preserving order.

## Adaptive Card Template
- Default template: `main.json` (Adaptive Card 1.3)
- Placeholders are replaced across all string fields:
  - `{{account}}`, `{{opportunity}}`, `{{amount}}`, `{{due}}`, `{{cta_url}}`
- After substitution the template is pruned:
  - Empty `TextBlock` nodes are removed
  - `FactSet` entries with blank values are removed (and the `FactSet` if empty)
  - `Action.OpenUrl` with blank `url` is removed
  - Empty containers/sets are dropped; the root `AdaptiveCard` is preserved
- Fallback markdown (for notifications or non‑card clients) is auto‑generated from the provided fields.

## Batching, Retries, and Logging
- Batching: sends up to `batch_size` recipients, then sleeps `batch_delay` seconds before the next batch.
- Retries: per‑recipient attempts up to `retry_count`, sleeping `retry_delay` seconds between attempts when failures occur or exceptions are raised.
- Logging: appends to `send_log.csv` (or `log_file`). Columns:
  - `timestamp_utc, email, status, attempts, http_status, message_id, error_preview`

## CLI Reference
```text
python3 send_webex_notifications.py [--settings settings.json] [--csv recipients.csv]
  [--token TOKEN] [--batch-size N] [--batch-delay SECONDS]
  [--retry-count N] [--retry-delay SECONDS] [--log-file PATH]
  [--account STR] [--opportunity STR] [--amount STR] [--due YYYY-MM-DD]
  [--cta-url URL] [--card-json PATH] [--dry-run]
```

## Development
- Style: Python 3, PEP 8, type hints; IO is kept in small helpers.
- Quick smoke test: run with `--dry-run`.
- Tests (optional): use `pytest` + request mocking; target 80%+ coverage for new code (batching, retries, CSV parsing).

## Security
- Keep tokens out of Git: use `.env` locally (see `.env_SAMPLE`) and export with your shell.
- Do not commit `recipients.csv` or `send_log.csv` to public repos.
- Start with a small CSV and `--dry-run` before wide sends.
- Respect rate limits via `--batch-size` and `--batch-delay`.

## Troubleshooting
- 401 Unauthorized: verify `WEBEX_BOT_TOKEN` or `--token` and bot permissions.
- 400 Bad Request: check recipient email addresses and card JSON validity.
- 429 Too Many Requests: increase `--batch-delay` or lower `--batch-size`.
- Network timeouts or intermittent failures: rely on built‑in retries or increase `--retry-delay`.

## Files
- `send_webex_notifications.py` — main CLI
- `main.json` — Adaptive Card template with placeholders
- `recipients.csv` — input emails
- `send_log.csv` — append‑only results log
- `settings.json` — optional overrides
- `.env_SAMPLE` — sample env file for local use
