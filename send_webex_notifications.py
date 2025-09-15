#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bulk-send 1:1 Webex messages with an Adaptive Card via a bot token.
- Reads recipients from a CSV (first column or 'email' header).
- Sends messages in batches of 10, pausing 5 seconds between batches.
- Retries failed sends up to 3 times with a 5-second wait between attempts.
- Logs results to a local CSV log (default: send_log.csv).

Usage:
  export WEBEX_BOT_TOKEN="xxx"
  # Recommended: configure values in settings.json and run with no args
  python3 send_webex_notifications.py

  # Optional: override settings.json with flags
  python3 send_webex_notifications.py --csv recipients.csv \
      --account "ACME Corp" \
      --opportunity "Q4 Expansion" \
      --amount "$50,000" \
      --due "2025-10-01" \
      --cta-url "https://example.crm.com/opportunities/ACME-Q4" \
      --card-json main.json

Docs used:
- Create message (attachments for Adaptive Cards) and payload format. 
- Buttons & Cards (only one card per message; include fallback markdown). 
- Rate limiting guidance and Retry-After header behavior.
"""

import os
import sys
import csv
import time
import json
import argparse
from pathlib import Path
from datetime import datetime, timezone

import requests


WEBEX_MESSAGES_URL = "https://webexapis.com/v1/messages"

# Defaults (used when not provided via CLI or settings.json)
DEFAULT_SETTINGS_PATH = "settings.json"
DEFAULT_CSV = "recipients.csv"
DEFAULT_BATCH_SIZE = 10
DEFAULT_BATCH_DELAY = 5.0
DEFAULT_RETRY_COUNT = 3
DEFAULT_RETRY_DELAY = 5.0
DEFAULT_LOG_FILE = "send_log.csv"
DEFAULT_ACCOUNT = "ACME Corp"
DEFAULT_OPPORTUNITY = "New Sales Opportunity"
DEFAULT_AMOUNT = "$50,000"
DEFAULT_CTA_URL = "https://example.crm.com/opportunities/ABC123"
DEFAULT_CARD_JSON = "main.json"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Bulk 1:1 Webex notifier with Adaptive Card (bot token)."
    )
    p.add_argument("--settings", default=DEFAULT_SETTINGS_PATH, help="Path to settings JSON (default: settings.json).")
    p.add_argument("--csv", help="Path to recipients CSV (overrides settings).")
    p.add_argument("--token", help="Webex bot token. Overrides env/settings; otherwise uses WEBEX_BOT_TOKEN env var.")
    p.add_argument("--batch-size", type=int, help="Messages per batch (overrides settings).")
    p.add_argument("--batch-delay", type=float, help="Seconds to pause between batches (overrides settings).")
    p.add_argument("--retry-count", type=int, help="Max attempts per recipient (overrides settings).")
    p.add_argument("--retry-delay", type=float, help="Seconds to wait between retries (overrides settings).")
    p.add_argument("--log-file", help="CSV log output path (overrides settings).")
    p.add_argument("--account", help="Account name to embed in the card (overrides settings).")
    p.add_argument("--opportunity", help="Opportunity title/name (overrides settings).")
    p.add_argument("--amount", help="Opportunity amount/value (overrides settings).")
    p.add_argument("--due", help="Due date string (e.g., 2025-10-01) (overrides settings).")
    p.add_argument("--cta-url", help="URL for the primary CTA button (overrides settings).")
    p.add_argument(
        "--card-json",
        help="Path to Adaptive Card JSON template to send (overrides settings).",
    )
    p.add_argument("--dry-run", action="store_true", help="Print what would be sent, but don't call the API.")
    return p.parse_args()


def load_emails(csv_path: str) -> list[str]:
    """Load recipient emails from a CSV (supports header 'email' or first column)."""
    emails: list[str] = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        # Peek first line to see if a header likely exists
        first_line = f.readline()
        f.seek(0)
        if "email" in first_line.lower():
            reader = csv.DictReader(f)
            for row in reader:
                email = (row.get("email") or row.get("Email") or row.get("EMAIL") or "").strip()
                if email and "@" in email:
                    emails.append(email)
        else:
            reader = csv.reader(f)
            for row in reader:
                if not row:
                    continue
                email = row[0].strip()
                if email and "@" in email and not email.lower().startswith("email"):
                    emails.append(email)

    # De-duplicate while preserving order
    seen = set()
    ordered = []
    for e in emails:
        if e not in seen:
            seen.add(e)
            ordered.append(e)
    return ordered


def chunked(items: list[str], size: int):
    for i in range(0, len(items), size):
        yield items[i : i + size]


def load_settings(path: str) -> dict:
    """Load optional settings JSON. Returns empty dict if not found.

    Supported keys (snake_case):
    - csv, log_file, batch_size, batch_delay, retry_count, retry_delay,
      account, opportunity, amount, due, cta_url, card_json, token
    """
    p = Path(path)
    if not p.exists():
        return {}
    try:
        with p.open("r", encoding="utf-8") as f:
            data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("settings JSON must be an object")
            return data
    except json.JSONDecodeError as e:
        raise SystemExit(f"ERROR: Invalid JSON in settings file {path}: {e}")
    except OSError as e:
        raise SystemExit(f"ERROR: Unable to read settings file {path}: {e}")


def _deep_replace_placeholders(value, variables: dict):
    """Recursively replace {{placeholders}} in strings within JSON-like structures.

    This allows a JSON template (e.g., main.json) to include tokens like
    {{account}}, {{opportunity}}, {{amount}}, {{due}}, {{cta_url}} which will
    be replaced with CLI-provided values.
    """
    if isinstance(value, str):
        out = value
        for k, v in variables.items():
            token = f"{{{{{k}}}}}"
            out = out.replace(token, v if v is not None else "")
        return out
    if isinstance(value, list):
        return [_deep_replace_placeholders(x, variables) for x in value]
    if isinstance(value, dict):
        return {k: _deep_replace_placeholders(v, variables) for k, v in value.items()}
    return value


def load_card_json(template_path: str, variables: dict) -> dict:
    """Load an Adaptive Card JSON file and apply placeholder substitution.

    - Reads JSON from `template_path`.
    - Performs simple string replacement for tokens like {{account}} across all
      string values in the JSON structure, using `variables`.
    - Returns a dict suitable to send as `attachments[0].content`.
    """
    try:
        with open(template_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except FileNotFoundError:
        raise SystemExit(f"ERROR: Card template not found: {template_path}")
    except json.JSONDecodeError as e:
        raise SystemExit(f"ERROR: Invalid JSON in card template {template_path}: {e}")

    # Apply simple token replacement across the loaded JSON structure
    rendered = _deep_replace_placeholders(raw, variables)

    def _is_blank(s: str | None) -> bool:
        return s is None or (isinstance(s, str) and s.strip() == "")

    def _prune(node, parent_type: str | None = None):
        """Remove empty elements produced by substitution (e.g., blank Due).

        - Drops TextBlock with empty text.
        - Drops FactSet facts where value is blank; removes FactSet if no facts.
        - Drops Action.OpenUrl when url is blank.
        - Removes empty containers/sets (Container.items, ColumnSet.columns, ActionSet.actions),
          but always preserves the AdaptiveCard root object.
        """
        if isinstance(node, dict):
            node_type = node.get("type")

            # Special-case pruning before recursing into children
            if node_type == "TextBlock" and _is_blank(node.get("text")):
                return None
            if node_type == "FactSet":
                facts = node.get("facts", [])
                pruned_facts = [f for f in facts if not _is_blank((f or {}).get("value"))]
                if not pruned_facts:
                    return None
                node = {**node, "facts": pruned_facts}
            if node_type == "Action.OpenUrl" and _is_blank(node.get("url")):
                return None

            # Recurse into dict children
            out = {}
            for k, v in node.items():
                if isinstance(v, (dict, list)):
                    pruned = _prune(v, node_type)
                    if pruned is None:
                        continue
                    out[k] = pruned
                else:
                    out[k] = v

            # Remove empty groupings (except for AdaptiveCard root)
            if node_type in ("Container", "Column") and len(out.get("items", [])) == 0:
                return None
            if node_type == "ColumnSet" and len(out.get("columns", [])) == 0:
                return None
            if node_type == "ActionSet" and len(out.get("actions", [])) == 0:
                return None
            # Keep AdaptiveCard even if body/actions become empty
            return out

        if isinstance(node, list):
            children = []
            for item in node:
                pruned = _prune(item, parent_type)
                if pruned is not None:
                    children.append(pruned)
            return children

        return node

    rendered = _prune(rendered)

    # Ensure minimal required fields exist
    if not isinstance(rendered, dict) or rendered.get("type") != "AdaptiveCard":
        raise SystemExit("ERROR: Card template must be an AdaptiveCard object with type 'AdaptiveCard'.")

    return rendered


def build_fallback_markdown(account: str, opportunity: str, amount: str, due: str | None) -> str:
    """Markdown fallback text shown in notifications or if card can't render."""
    base = f"You have a new **sales opportunity** assigned: **{account} — {opportunity} ({amount})**."
    if due:
        base += f" Due: **{due}**."
    base += " See the attached Adaptive Card for details and actions."
    return base


def send_message_to_email(
    bot_token: str, to_email: str, markdown: str, adaptive_card: dict, timeout: int = 30
) -> requests.Response:
    headers = {
        "Authorization": f"Bearer {bot_token}",
        "Content-Type": "application/json",
    }
    payload = {
        "toPersonEmail": to_email,
        "markdown": markdown,  # Fallback text is required when sending a card
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": adaptive_card,
            }
        ],
    }
    return requests.post(WEBEX_MESSAGES_URL, headers=headers, json=payload, timeout=timeout)


def ensure_log_writer(log_path: str):
    first_write = not Path(log_path).exists()
    f = open(log_path, "a", newline="", encoding="utf-8")
    writer = csv.writer(f)
    if first_write:
        writer.writerow(["timestamp_utc", "email", "status", "attempts", "http_status", "message_id", "error_preview"])
    return f, writer


def main():
    args = parse_args()

    # Load settings (optional), then merge with CLI and built-in defaults
    settings = load_settings(args.settings or DEFAULT_SETTINGS_PATH)

    # Token from CLI, then ENV, then settings (avoid keeping secrets in files)
    token = args.token or os.getenv("WEBEX_BOT_TOKEN") or settings.get("token")
    if not token:
        print("ERROR: Provide a bot token via --token, WEBEX_BOT_TOKEN env var, or settings.json.", file=sys.stderr)
        sys.exit(1)

    # Resolve configuration with precedence: CLI value -> settings.json -> default
    csv_path = args.csv or settings.get("csv") or DEFAULT_CSV
    log_file = args.log_file or settings.get("log_file") or DEFAULT_LOG_FILE
    batch_size = args.batch_size if args.batch_size is not None else settings.get("batch_size", DEFAULT_BATCH_SIZE)
    batch_delay = args.batch_delay if args.batch_delay is not None else settings.get("batch_delay", DEFAULT_BATCH_DELAY)
    retry_count = args.retry_count if args.retry_count is not None else settings.get("retry_count", DEFAULT_RETRY_COUNT)
    retry_delay = args.retry_delay if args.retry_delay is not None else settings.get("retry_delay", DEFAULT_RETRY_DELAY)
    account = args.account or settings.get("account") or DEFAULT_ACCOUNT
    opportunity = args.opportunity or settings.get("opportunity") or DEFAULT_OPPORTUNITY
    amount = args.amount or settings.get("amount") or DEFAULT_AMOUNT
    due = args.due if args.due is not None else settings.get("due")
    cta_url = args.cta_url or settings.get("cta_url") or DEFAULT_CTA_URL
    card_json = args.card_json or settings.get("card_json") or DEFAULT_CARD_JSON

    if not Path(csv_path).exists():
        print(f"ERROR: CSV file not found: {csv_path}", file=sys.stderr)
        sys.exit(1)

    emails = load_emails(csv_path)
    if not emails:
        print("No valid recipient emails found in CSV.", file=sys.stderr)
        sys.exit(1)

    print(f"Loaded {len(emails)} recipient(s).")

    # Build shared message template from JSON file with placeholder substitution
    variables = {
        "account": account,
        "opportunity": opportunity,
        "amount": amount,
        "due": (due or ""),
        "cta_url": cta_url,
    }
    card = load_card_json(card_json, variables)
    fallback_md = build_fallback_markdown(account, opportunity, amount, due)

    # Prepare log
    log_file_handle, log_writer = ensure_log_writer(log_file)

    total_sent = 0
    total_failed = 0

    try:
        batch_index = 0
        for batch in chunked(emails, batch_size):
            batch_index += 1
            print(f"\n=== Batch {batch_index}: sending {len(batch)} message(s) ===")
            for to_email in batch:
                attempts = 0
                last_status = None
                message_id = ""
                error_preview = ""
                ok = False

                for attempt in range(1, retry_count + 1):
                    attempts = attempt
                    if args.dry_run:
                        ok = True
                        last_status = 200
                        message_id = "(dry-run)"
                        break

                    try:
                        resp = send_message_to_email(token, to_email, fallback_md, card, timeout=30)
                        last_status = resp.status_code
                        if resp.status_code in (200, 201):  # 200 OK usually returned
                            data = {}
                            try:
                                data = resp.json()
                            except Exception:
                                pass
                            message_id = data.get("id", "")
                            ok = True
                            print(f"[OK] {to_email} (attempt {attempts})  id={message_id}")
                            break
                        else:
                            # Capture a short preview of the error body
                            error_preview = (resp.text or "")[:300].replace("\n", " ")
                            print(f"[WARN] Attempt {attempts} for {to_email} failed: {last_status} {error_preview}")
                            if attempts < retry_count:
                                time.sleep(retry_delay)
                    except requests.RequestException as e:
                        error_preview = str(e)[:300].replace("\n", " ")
                        print(f"[WARN] Attempt {attempts} for {to_email} raised exception: {error_preview}")
                        if attempts < retry_count:
                            time.sleep(retry_delay)

                # Log result
                ts = datetime.now(timezone.utc).isoformat()
                if ok:
                    total_sent += 1
                    log_writer.writerow([ts, to_email, "sent", attempts, last_status, message_id, ""])
                else:
                    total_failed += 1
                    log_writer.writerow([ts, to_email, "failed", attempts, last_status, message_id, error_preview])

            # Inter-batch delay (skip after final batch)
            if (batch_index * batch_size) < len(emails):
                print(f"Batch {batch_index} complete. Pausing {batch_delay} sec to respect rate limits...")
                time.sleep(batch_delay)

    finally:
        log_file_handle.flush()
        log_file_handle.close()

    print("\n=== Summary ===")
    print(f"Sent:   {total_sent}")
    print(f"Failed: {total_failed}")
    print(f"Log written to: {log_file}")


if __name__ == "__main__":
    main()
