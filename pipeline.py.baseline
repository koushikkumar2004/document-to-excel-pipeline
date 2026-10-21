

import argparse

import configparser

import json

import re

import subprocess

import sys

import time

from datetime import date, datetime

from pathlib import Path

import requests

from openpyxl import Workbook

from pypdf import PdfReader

# ============================================================

# PATHS AND CONFIGURATION

# ============================================================

BASE_DIR = Path(__file__).resolve().parent

CONFIG_PATH = BASE_DIR / "config.ini"

if hasattr(sys.stdout, "reconfigure"):

    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

if hasattr(sys.stderr, "reconfigure"):

    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

config = configparser.ConfigParser()

try:

    with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:

        config.read_file(f)

except (OSError, configparser.Error) as exc:

    raise SystemExit(f"ERROR loading config.ini: {exc}")

def setting(section, key, default=None):

    if section not in config or key not in config[section]:

        if default is not None:

            return default

        raise RuntimeError(f"Missing [{section}] {key} in config.ini")

    value = config[section][key].strip()

    if not value and default is None:

        raise RuntimeError(f"Empty setting [{section}] {key}")

    return value if value else default

def integer_setting(section, key, default):

    try:

        return int(setting(section, key, str(default)))

    except ValueError:

        raise RuntimeError(f"[{section}] {key} must be an integer")

def float_setting(section, key, default):

    try:

        return float(setting(section, key, str(default)))

    except ValueError:

        raise RuntimeError(f"[{section}] {key} must be numeric")

BASE_URL = setting("lm_studio", "base_url").rstrip("/")

MODEL = setting("lm_studio", "model")

TOKEN = setting("lm_studio", "token", "")

TIMEOUT = integer_setting("lm_studio", "timeout_seconds", 600)

TEMPERATURE = float_setting("lm_studio", "temperature", 0)

MAX_TOKENS = integer_setting("lm_studio", "max_output_tokens", 4096)

BATCH_SIZE = integer_setting("pipeline", "batch_size", 1)

SPARSE_THRESHOLD = integer_setting(

    "pipeline", "sparse_page_threshold", 20

)

EXTRA_COUNT = integer_setting("pipeline", "extra_slot_count", 10)

PDFTOTEXT = Path(

    setting("pipeline", "pdftotext_path", "bin/pdftotext.exe")

)

if not PDFTOTEXT.is_absolute():

    PDFTOTEXT = BASE_DIR / PDFTOTEXT

if BATCH_SIZE < 1 or TIMEOUT < 1 or MAX_TOKENS < 1:

    raise RuntimeError("Batch size, timeout and max tokens must be positive.")

if EXTRA_COUNT < 0:

    raise RuntimeError("extra_slot_count cannot be negative.")

API_URL = BASE_URL + "/chat/completions"

# ============================================================

# FIXED SCHEMA

# ============================================================

HEADER_COLUMNS = [

    "document_type",

    "document_number",

    "document_date",

    "seller_name",

    "seller_address",

    "seller_tax_id",

    "buyer_name",

    "buyer_address",

    "buyer_tax_id",

    "reference_number",

    "due_date",

    "currency",

    "subtotal_amount",

    "discount_amount",

    "tax_amount",

    "total_amount",

    "remarks",

]

LINE_ITEM_COLUMNS = [

    "item_code",

    "item_description",

    "quantity",

    "unit",

    "rate",

    "line_amount",

]

EXTRA_COLUMNS = [

    f"extra_{i}" for i in range(1, EXTRA_COUNT + 1)

]

FIXED_COLUMNS = HEADER_COLUMNS + LINE_ITEM_COLUMNS + EXTRA_COLUMNS

ALLOWED_COLUMNS = set(FIXED_COLUMNS)

FINAL_COLUMNS = FIXED_COLUMNS + ["source_page"]

FIXED_TYPES = {

    "document_date": "date",

    "due_date": "date",

    "subtotal_amount": "numeric",

    "discount_amount": "numeric",

    "tax_amount": "numeric",

    "total_amount": "numeric",

    "quantity": "numeric",

    "rate": "numeric",

    "line_amount": "numeric",

}

# ============================================================

# RUN REPORTS

# ============================================================

INVENTED_KEYS = []

TYPE_WARNINGS = []

DROPPED_COLUMNS = []

def log(message=""):

    print(message, flush=True)

# ============================================================

# TEXT EXTRACTION

# ============================================================

def extract_pdf(pdf_path):

    if not PDFTOTEXT.exists():

        raise RuntimeError(f"pdftotext executable not found: {PDFTOTEXT}")

    reader = PdfReader(str(pdf_path))

    if reader.is_encrypted:

        raise RuntimeError("PDF is encrypted or password protected.")

    page_count = len(reader.pages)

    result = subprocess.run(

        [

            str(PDFTOTEXT),

            "-layout",

            "-enc", "UTF-8",

            str(pdf_path),

            "-"

        ],

        capture_output=True,

        text=True,

        encoding="utf-8",

        errors="replace",

        check=False,

    )

    if result.returncode != 0:

        raise RuntimeError(f"pdftotext failed: {result.stderr}")

    page_texts = result.stdout.split("\f")

    if page_texts and not page_texts[-1].strip():

        page_texts.pop()

    if len(page_texts) != page_count:

        raise RuntimeError(

            f"PDF page count mismatch: expected {page_count}, "

            f"received {len(page_texts)}."

        )

    parts = []

    for number, text in enumerate(page_texts, start=1):

        count = len("".join(text.split()))

        if count < SPARSE_THRESHOLD:

            log(f"WARNING: Page {number} has only {count} non-whitespace characters.")

        parts.append(f"--- Page {number} ---\n{text.strip()}")

    return "\n\n".join(parts), page_count

def extract_txt(txt_path):

    text = txt_path.read_text(encoding="utf-8-sig", errors="replace")

    if len("".join(text.split())) < SPARSE_THRESHOLD:

        log("WARNING: TXT file contains very little text.")

    return text, None

def split_pages(text, expected_count=None):

    # PDF extraction includes explicit page markers.

    if expected_count is not None:

        matches = list(

            re.finditer(r"(?m)^--- Page (\d+) ---\s*$", text)

        )

        if len(matches) != expected_count:

            raise RuntimeError(

                f"Could not identify all page markers: expected "

                f"{expected_count}, found {len(matches)}."

            )

        pages = []

        for i, match in enumerate(matches):

            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)

            pages.append({

                "page": int(match.group(1)),

                "text": text[match.end():end].strip()

            })

        return pages

    # A TXT file is treated as one page unless it has PDF-style markers.

    matches = list(

        re.finditer(r"(?m)^--- Page (\d+) ---\s*$", text)

    )

    if matches:

        pages = []

        for i, match in enumerate(matches):

            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)

            pages.append({

                "page": int(match.group(1)),

                "text": text[match.end():end].strip()

            })

        return pages

    return [{"page": 1, "text": text.strip()}]

# ============================================================

# EXTRACTION PROMPT

# ============================================================

def build_prompt(batch, used_keys):

    document_text = "\n\n".join(

        f"--- Page {p['page']} ---\n{p['text']}"

        for p in batch

    )

    columns = "\n".join(FIXED_COLUMNS)

    continuation = (

        "\nPreviously observed approved keys: "

        + ", ".join(sorted(used_keys))

        + "\nReuse existing keys for the same concepts.\n"

        if used_keys else ""

    )

    return f"""

You are a precise document information extraction system.

Extract only facts supported by the supplied document.

Use ONLY these fixed column names:

{columns}

GENERAL RULES

1. Never invent, guess, calculate, or fill missing values.

2. Represent absent or blank source values as "".

3. Preserve the distinction between source columns.

4. Do not combine values from different table rows.

5. Do not combine prices from different items.

6. Do not create additional column names.

7. Do not put a currency symbol alone into a monetary field.

8. Preserve printed values and their meaning.

9. Do not treat subtotal, tax, shipping, or grand-total rows

   as individual items.

LINE-ITEM TABLE RULES — CRITICAL

1. Every distinct source-table item row must become exactly

   one object in "line_items".

2. If the source has four items, return four line-item objects.

3. Read each table row horizontally. Keep its values together.

4. Never merge two or more source rows into one item.

5. Do not put several prices into one field.

6. Map a product/service name to "item_code" when that is

   the closest available named field.

7. Map the explanatory text to "item_description".

8. Map quantity only if the source explicitly supplies it.

   A blank quantity must remain blank, not 1.

9. Map units such as hour, hours, kg, or pieces to "unit".

10. Map a printed unit price to "rate".

11. Map a printed item total to "line_amount".

12. Do not copy a rate into line_amount if the line total is blank.

13. Do not calculate a missing line total.

14. If a total cell contains only "$" or another currency symbol,

    return "" for line_amount.

15. If a cell is unclear, do not guess its value.

RATES, UNITS AND CURRENCY

- For a value such as "$50.00/hour", extract the numeric rate

  as 50.00 and the unit as "hour" when the source supports it.

- Do not place "$" alone into rate or line_amount.

- Use currency for the currency code/name if identifiable.

- Keep meaningful units and descriptions.

DISCOUNTS

- discount_amount is only for a printed monetary discount.

- Do not put a percentage such as 20% into discount_amount.

- There is no dedicated discount percentage column in this schema.

  Do not invent one or misrepresent the percentage.

HEADER RULES — LABELS MUST MATCH

- document_number is the value printed beside labels such as "Invoice #",
  "Invoice No.", "Document Number", or equivalent.
- reference_number is ONLY a separate PO/reference/order number printed beside
  a reference label. Never copy document_number into reference_number. If no
  separate reference value is printed, use "".
- document_date is the value beside "Invoice Date" or "Document Date".
- due_date is ONLY a value beside "Due Date" / "Payment Due". Never copy
  document_date into due_date. If no due date is printed, use "".
- subtotal_amount must match the amount printed beside "Subtotal".
- total_amount should match the final invoice total / "Balance Due" / "Grand
  Total" when that label is present. Do not substitute "Amount Due" or a
  remittance amount if a separate balance/grand total is printed.
- Never calculate or derive an amount from quantity, rate, discount, or other
  values. Use the amount printed in that item's own Total column.
- Keep seller and buyer details separate where identifiable.

EXTRA SLOTS

- Use extra_1 through extra_{EXTRA_COUNT} only for facts not

  covered by named columns.

- Reuse a slot only for the same concept.

- Do not use extra slots instead of a suitable named field.

RECORD RULES

- Keep unrelated documents separate.

- A page that continues a table should not be treated as a new

  unrelated document.

- The "page" value must be a supplied page number.

{continuation}

Return ONLY valid JSON in this structure:

{{

  "records": [

    {{

      "page": 1,

      "header": {{}},

      "line_items": []

    }}

  ],

  "column_types": {{}}

}}

Do not add Markdown fences or explanations.

DOCUMENT:

{document_text}

"""

# ============================================================

# LM STUDIO AND JSON PARSING

# ============================================================

def request_model(prompt):

    headers = {"Content-Type": "application/json"}

    if TOKEN:

        headers["Authorization"] = f"Bearer {TOKEN}"

    payload = {

        "model": MODEL,

        "messages": [

            {

                "role": "system",

                "content": "Extract document facts and return valid JSON only."

            },

            {"role": "user", "content": prompt}

        ],

        "temperature": TEMPERATURE,

        "max_tokens": MAX_TOKENS,

    }

    try:

        response = requests.post(

            API_URL,

            headers=headers,

            json=payload,

            timeout=TIMEOUT,

        )

    except requests.exceptions.Timeout:

        raise RuntimeError(f"LM Studio timed out after {TIMEOUT} seconds.")

    except requests.exceptions.ConnectionError:

        raise RuntimeError(

            "Cannot connect to LM Studio. Check that the local server is running."

        )

    except requests.exceptions.RequestException as exc:

        raise RuntimeError(f"LM Studio request failed: {exc}")

    if response.status_code == 401:

        raise RuntimeError("HTTP 401: Check the LM Studio token.")

    if response.status_code != 200:

        raise RuntimeError(

            f"LM Studio HTTP {response.status_code}: {response.text[:1000]}"

        )

    try:

        return response.json()

    except ValueError:

        raise RuntimeError("LM Studio returned invalid HTTP JSON.")

def parse_response(raw):

    choices = raw.get("choices", [])

    if not choices:

        raise ValueError("Response has no choices.")

    choice = choices[0]

    message = choice.get("message", {})

    content = message.get("content", "")

    if not isinstance(content, str) or not content.strip():

        raise ValueError(

            f"Empty model content; finish_reason={choice.get('finish_reason')}"

        )

    content = re.sub(

        r"<think>.*?</think>",

        "",

        content,

        flags=re.DOTALL | re.IGNORECASE

    ).strip()

    content = re.sub(

        r"^\s*```(?:json)?\s*|\s*```\s*$",

        "",

        content,

        flags=re.IGNORECASE

    ).strip()

    try:

        parsed = json.loads(content)

    except json.JSONDecodeError:

        start, end = content.find("{"), content.rfind("}")

        if start < 0 or end <= start:

            raise ValueError("No JSON object found in model content.")

        parsed = json.loads(content[start:end + 1])

    if not isinstance(parsed, dict):

        raise ValueError("Model output must be a JSON object.")

    if not isinstance(parsed.get("records"), list):

        raise ValueError("Model JSON must contain a records list.")

    for record in parsed["records"]:

        if not isinstance(record, dict):

            raise ValueError("Each record must be an object.")

        if not isinstance(record.get("header", {}), dict):

            raise ValueError("Record header must be an object.")

        if not isinstance(record.get("line_items", []), list):

            raise ValueError("line_items must be a list.")

    return parsed

def has_usable_extraction(parsed):

    """Return True only when at least one non-empty extracted fact exists."""

    records = parsed.get("records", [])

    if not isinstance(records, list) or not records:

        return False

    for record in records:

        if not isinstance(record, dict):

            continue

        header = record.get("header", {})

        items = record.get("line_items", [])

        if isinstance(header, dict):

            for value in header.values():

                if value not in ("", None, [], {}):

                    return True

        if isinstance(items, list):

            for item in items:

                if isinstance(item, dict):

                    for value in item.values():

                        if value not in ("", None, [], {}):

                            return True

    return False

def build_recovery_prompt(batch, used_keys):

    """Short, direct second attempt used when the model returns empty records."""

    document_text = "\n\n".join(

        f"--- Page {page['page']} ---\n{page['text']}"

        for page in batch

    )

    columns = "\n".join(FIXED_COLUMNS)

    previously_used = (

        "Use these same keys for matching concepts where appropriate: "

        + ", ".join(sorted(used_keys))

        + "\n"

        if used_keys else ""

    )

    return f"""

Extract invoice/document data from the text below. The previous attempt returned

an empty record, but you must inspect the visible text and extract the facts that

are actually present.

Use only these output field names:

{columns}

Rules:

- Return JSON only, with top-level keys "records" and "column_types".

- Each record must have "page", "header", and "line_items".

- Use the supplied page number.

- Put document-level facts in "header".

- Put EACH distinct product/service table row into its own object in "line_items".

- Keep each product name and its own description/rate together.

- Use item_code for a product/service name when no dedicated item-name column exists.

- Use item_description for its descriptive text.

- Put a printed unit price in rate. For "$50.00/hour", rate is "50.00"
  and unit is "hour" (the unit suffix of the printed unit price).
- The unit field should contain ONLY a concise measurement unit such as hour,
  day, kg, item, or piece. Do not copy description phrases into unit.
- item_description must preserve descriptive text, e.g. "8 Hours of Website
  Development"; do not use the description as the unit.
- Use the printed item Total exactly. Never calculate line_amount from quantity
  and rate, even when a discount percentage appears on the row. If the rate and
  quantity imply 400 but the printed row total says 320, extract the printed
  320, not a calculated 400.
- For headers, match the value to the adjacent printed label. "Amount Due" and
  "Balance Due" are distinct labels and may have different values.
- If quantity or line total is blank in the source, return an empty string for it.
- Never calculate totals, invent values, or combine different products/prices.

- Leave unknown fields as empty strings. Do not return an empty record if facts

  are visible in the supplied text.

- Do not create keys outside the allowed field names.

- Use extra_1 through extra_{EXTRA_COUNT} only for facts with no suitable named field.

{previously_used}

Expected JSON shape (fill it with extracted facts; do not copy empty values blindly):

{{"records":[{{"page":1,"header":{{"document_type":"invoice"}},"line_items":[{{"item_code":"example product","item_description":"example description","rate":"12.50"}}]}}],"column_types":{{}}}}

DOCUMENT TEXT:

{document_text}

""".strip()

def _normalized_numeric(value):
    """Return the first numeric token in a value, normalized as text."""
    if value is None or value == "":
        return None
    match = re.search(r"[-+]?\d[\d,]*(?:.\d+)?", str(value))
    if not match:
        return None
    return match.group(0).replace(",", "")

def extraction_quality_issues(parsed, batch):
    """Detect merged rows and obvious values that contradict printed evidence."""
    issues = []
    document_text = "\n".join(page.get("text", "") for page in batch)
    normalized_text = re.sub(r"(?<=\d),(?=\d{3}\b)", "", document_text)
    text_numbers = {
        token.replace(",", "")
        for token in re.findall(r"(?<![A-Za-z])[-+]?\d[\d,]*(?:.\d+)?", normalized_text)
    }

    total_label_pattern = re.compile(
        r"\b(subtotal|grand total|total due|balance due|amount due|amount enclosed|"
        r"tax|shipping|miscellaneous|discount total)\b",
        re.IGNORECASE,
    )
    price_rows = [
        line for line in document_text.splitlines()
        if not total_label_pattern.search(line)
        and re.search(r"[$€£₹]\s*\d[\d,]*(?:.\d{2})?", line)
    ]

    for record_index, record in enumerate(parsed.get("records", []), start=1):
        items = record.get("line_items", [])
        if not isinstance(items, list):
            items = []

        for item_index, item in enumerate(items, start=1):
            if not isinstance(item, dict):
                continue
            for field in ("quantity", "rate", "line_amount"):
                value = item.get(field, "")
                if isinstance(value, str) and value.strip():
                    if re.search(r",\s*(?:[$€£₹]?\s*)[-+]?\d", value.strip()):
                        issues.append(
                            f"record {record_index}, item {item_index}: {field} contains multiple comma-separated values ({value!r})"
                        )

            amount = _normalized_numeric(item.get("line_amount", ""))
            if amount and amount not in text_numbers:
                issues.append(
                    f"record {record_index}, item {item_index}: line_amount {item.get('line_amount')!r} has no matching numeric value in source text"
                )

        if len(price_rows) >= 2 and len(items) == 1:
            issues.append(
                f"document text has {len(price_rows)} price-bearing lines but the model returned only one line item"
            )

        header = record.get("header", {})
        if isinstance(header, dict):
            doc_num = str(header.get("document_number", "")).strip()
            ref_num = str(header.get("reference_number", "")).strip()
            if doc_num and ref_num and doc_num == ref_num:
                has_reference_label = re.search(
                    r"\b(?:PO\s*(?:#|No.?)?|purchase order|reference(?:\s*(?:#|no.?))?|your order #|our order #)\b",
                    document_text,
                    re.IGNORECASE,
                )
                if not has_reference_label:
                    issues.append(
                        "reference_number duplicates document_number without a clearly labeled reference field"
                    )

            doc_date = str(header.get("document_date", "")).strip()
            due_date = str(header.get("due_date", "")).strip()
            if doc_date and due_date and doc_date == due_date and not re.search(
                r"\b(?:due date|payment due)\b", document_text, re.IGNORECASE
            ):
                issues.append(
                    "due_date duplicates document_date but the source has no explicit due-date label"
                )

            total_value = _normalized_numeric(header.get("total_amount", ""))
            balance_match = re.search(
                r"\b(?:balance due|grand total|invoice total|total due)\b[^\r\n$€£₹]{0,30}[$€£₹]?\s*([\d,]+(?:.\d{2})?)",
                document_text,
                re.IGNORECASE,
            )
            if balance_match and total_value:
                printed_total = balance_match.group(1).replace(",", "")
                if total_value != printed_total:
                    issues.append(
                        f"total_amount {header.get('total_amount')!r} does not match the printed Balance Due/Grand Total {printed_total}"
                    )

    return list(dict.fromkeys(issues))

def build_table_recovery_prompt(batch, used_keys, prior_issues):
    """Third attempt focused on keeping source table rows separate."""
    prompt = build_recovery_prompt(batch, used_keys)
    issue_text = "\n".join(f"- {issue}" for issue in prior_issues)
    return prompt + f"""

TABLE-STRUCTURE RECOVERY — REQUIRED
The previous response failed structural/semantic validation:
{issue_text}

Correct each listed issue by re-reading the exact source labels and table cells.
Do not repeat a value merely because it was present in the previous response.
Re-read the source table row by row. Do not return the same combined item.
For each physical product/service row, create one separate object in line_items.
Never place comma-separated lists of different quantities, rates, item names, or
line amounts into one item's fields. Do not put document/header facts inside
line_items. Put invoice number/date/seller/buyer/subtotal/balance in header.
A continuation line belongs to the closest relevant item description and is not
a new item unless it has its own item row.
For any field you cannot confidently identify, return an empty string.
Check every numeric amount against the actual printed cell/label before returning:
do not calculate a line amount from rate and quantity; discounts may make printed
totals differ from multiplication. Use "Balance Due"/"Grand Total" for total_amount
when available, not a different "Amount Due" field. Do not duplicate document_number
as reference_number or document_date as due_date. Keep unit to a concise unit token.
Return JSON only, following the required schema.
""".strip()

def process_batch(batch, batch_number, used_keys):
    first_prompt = build_prompt(batch, used_keys)
    recovery_prompt = build_recovery_prompt(batch, used_keys)
    last_error = None
    prior_issues = []

    for attempt in range(1, 4):
        if attempt == 1:
            prompt = first_prompt
        elif attempt == 2:
            prompt = recovery_prompt
        else:
            prompt = build_table_recovery_prompt(batch, used_keys, prior_issues)

        log(f"  LM Studio attempt {attempt}/3...")
        raw = request_model(prompt)
        try:
            parsed = parse_response(raw)
        except (ValueError, json.JSONDecodeError) as exc:
            last_error = f"invalid model JSON: {exc}"
            log(f"  {last_error}")
            if attempt < 3:
                log("  Retrying with a stronger recovery prompt...")
            continue

        if not has_usable_extraction(parsed):
            last_error = "model returned valid JSON but extracted no non-empty facts"
            prior_issues = [last_error]
            log(f"  WARNING: {last_error}.")
            if attempt < 3:
                log("  Retrying with a stronger recovery prompt...")
            continue

        issues = extraction_quality_issues(parsed, batch)
        if issues:
            prior_issues = issues
            last_error = "extraction failed structural validation: " + "; ".join(issues)
            log("  WARNING: " + last_error)
            if attempt < 3:
                log("  Retrying because the table structure appears incorrect...")
            continue

        return raw, parsed, prompt

    raise RuntimeError(
        f"Batch {batch_number}: extraction failed after 3 attempts: {last_error}. "
        "No Excel file was written for this run. Inspect the extracted text and "
        "response JSON, then try a different local model or revise the prompt."
    )

# ============================================================
# VALIDATE KEYS AND PRESERVE VALUES
# ============================================================

STRUCTURAL_KEYS = {
    "records", "page", "header", "line_items", "column_types"
}

def inspect_fields(obj, batch_number, location="root"):

    if isinstance(obj, dict):

        for key, value in obj.items():

            if key not in ALLOWED_COLUMNS and key not in STRUCTURAL_KEYS:

                INVENTED_KEYS.append({

                    "batch": batch_number,

                    "key": key,

                    "location": location,

                })

                log(f"WARNING: Unexpected key '{key}' at {location}.")

            inspect_fields(value, batch_number, f"{location}.{key}")

    elif isinstance(obj, list):

        for index, item in enumerate(obj):

            inspect_fields(item, batch_number, f"{location}[{index}]")

def collect_used_keys(records):

    keys = set()

    for record in records:

        header = record.get("header", {})

        items = record.get("line_items", [])

        if isinstance(header, dict):

            keys.update(k for k in header if k in ALLOWED_COLUMNS)

        if isinstance(items, list):

            for item in items:

                if isinstance(item, dict):

                    keys.update(k for k in item if k in ALLOWED_COLUMNS)

    return keys

def first_nonempty_values(records):

    examples = {}

    for record in records:

        for section in (

            record.get("header", {}),

            *(record.get("line_items", []) or [])

        ):

            if not isinstance(section, dict):

                continue

            for key, value in section.items():

                if key in ALLOWED_COLUMNS and value not in ("", None):

                    examples.setdefault(key, str(value)[:100])

    return examples

# ============================================================

# FLATTEN RECORDS

# ============================================================

def merge_row(header, item, page):

    row = {}

    for key in FIXED_COLUMNS:

        value = ""

        if isinstance(header, dict):

            value = header.get(key, "")

        if isinstance(item, dict) and key in item:

            item_value = item.get(key, "")

            if item_value not in ("", None):

                value = item_value

        row[key] = value if value is not None else ""

    row["source_page"] = page

    return row

def flatten_fixed(records):

    rows = []

    for record in records:

        header = record.get("header", {})

        items = record.get("line_items", [])

        page = record.get("page", "")

        if not isinstance(header, dict):

            header = {}

        if not isinstance(items, list):

            items = []

        if items:

            for item in items:

                rows.append(merge_row(header, item, page))

        else:

            rows.append(merge_row(header, {}, page))

    return rows

def flatten_raw(records):

    rows = []

    discovered_columns = []

    def add_column(key):

        if key not in discovered_columns:

            discovered_columns.append(key)

    for record in records:

        header = record.get("header", {})

        items = record.get("line_items", [])

        page = record.get("page", "")

        if not isinstance(header, dict):

            header = {}

        if not isinstance(items, list):

            items = []

        item_list = items if items else [{}]

        for item in item_list:

            if not isinstance(item, dict):

                item = {}

            row = {}

            for section in (header, item):

                for key, value in section.items():

                    add_column(key)

                    if key not in row or row[key] in ("", None):

                        row[key] = value

            add_column("source_page")

            row["source_page"] = page

            rows.append(row)

    return rows, discovered_columns

# ============================================================

# TYPE CONVERSION FOR FINAL EXCEL ONLY

# ============================================================

def convert_number(value, column, excel_row):
    """Convert one printed numeric value, preserving unsupported originals."""
    if value is None or value == "":
        return ""

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()

    # Unit prices may include a unit suffix, e.g. "$50.00/hour".
    # The unit itself is handled by the extraction model in the unit column.
    if column == "rate":
        text = re.sub(
            r"\s*/\s*(?:hours?|hrs?|kilograms?|kg|pieces?|units?)\s*$",
            "",
            text,
            flags=re.IGNORECASE,
        ).strip()

    match = re.fullmatch(
        r"(?:INR|Rs.?|₹|\$|€|£)?\s*"
        r"([-+]?\d[\d,]*(?:.\d+)?)\s*(%)?",
        text,
        flags=re.IGNORECASE,
    )

    if not match:
        TYPE_WARNINGS.append({
            "column": column,
            "row": excel_row,
            "value": value,
            "expected": "numeric",
        })
        log(f"WARNING: Cannot convert {value!r} in {column}, Excel row {excel_row}.")
        return value

    number_text = match.group(1).replace(",", "")
    try:
        number = float(number_text)
        if match.group(2) == "%":
            number /= 100
        return int(number) if number.is_integer() else number
    except ValueError:
        TYPE_WARNINGS.append({
            "column": column,
            "row": excel_row,
            "value": value,
            "expected": "numeric",
        })
        log(f"WARNING: Cannot convert {value!r} in {column}, Excel row {excel_row}.")
        return value

def convert_date(value, column, excel_row):

    if value is None or value == "":

        return ""

    if isinstance(value, (date, datetime)):

        return value.date() if isinstance(value, datetime) else value

    text = str(value).strip()

    formats = [

        "%Y-%m-%d",

        "%Y-%m-%d %H:%M:%S",

        "%d/%m/%Y",

        "%d-%m-%Y",

        "%d.%m.%Y",

        "%d/%m/%y",

        "%d-%m-%y",

        "%d %B %Y",

        "%d %b %Y",

        "%m/%d/%Y",

    ]

    for fmt in formats:

        try:

            return datetime.strptime(text, fmt).date()

        except ValueError:

            pass

    TYPE_WARNINGS.append({

        "column": column,

        "row": excel_row,

        "value": value,

        "expected": "date",

    })

    log(f"WARNING: Cannot convert {value!r} in {column}, Excel row {excel_row}.")

    return value

def convert_final_value(column, value, excel_row, model_types):

    expected = FIXED_TYPES.get(column, model_types.get(column, "text"))

    if expected == "numeric":

        return convert_number(value, column, excel_row)

    if expected == "date":

        return convert_date(value, column, excel_row)

    return value

# ============================================================

# EXCEL WRITERS

# ============================================================

def save_workbook(rows, columns, output_path, convert=False, model_types=None):

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Extracted Data"

    for col_number, column in enumerate(columns, start=1):

        sheet.cell(row=1, column=col_number, value=column)

    for row_number, row in enumerate(rows, start=2):

        for col_number, column in enumerate(columns, start=1):

            value = row.get(column, "")

            if convert:

                value = convert_final_value(

                    column, value, row_number, model_types or {}

                )

            cell = sheet.cell(

                row=row_number,

                column=col_number,

                value=value if value is not None else ""

            )

            if column in ("document_date", "due_date") and isinstance(

                value, (date, datetime)

            ):

                cell.number_format = "yyyy-mm-dd"

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = sheet.dimensions

    workbook.save(output_path)

# ============================================================

# DEBUG OUTPUTS

# ============================================================

def write_json(path, value):

    path.write_text(

        json.dumps(value, indent=2, ensure_ascii=False, default=str),

        encoding="utf-8"

    )

# ============================================================

# MAIN

# ============================================================

def main():

    parser = argparse.ArgumentParser(

        description="Extract facts from a text-based PDF or TXT into Excel."

    )

    parser.add_argument("input_file", help="Input PDF or TXT file")

    parser.add_argument(

        "--output",

        help="Optional final Excel path (default: next to input, same base name)"

    )

    args = parser.parse_args()

    input_path = Path(args.input_file).resolve()

    if not input_path.is_file():

        raise RuntimeError(f"Input file not found: {input_path}")

    extension = input_path.suffix.lower()

    if extension not in (".pdf", ".txt"):

        raise RuntimeError("Only PDF and TXT inputs are supported.")

    final_path = (

        Path(args.output).resolve()

        if args.output

        else input_path.with_suffix(".xlsx")

    )

    raw_path = final_path.with_name(final_path.stem + "_raw.xlsx")

    extracted_path = input_path.with_name(input_path.stem + "_extracted.txt")

    response_path = input_path.with_name(input_path.stem + "_response.json")

    started = time.time()

    log("=" * 64)

    log("DOCUMENT-TO-EXCEL PIPELINE")

    log("=" * 64)

    log(f"Start time: {datetime.now():%Y-%m-%d %H:%M:%S}")

    log(f"Input: {input_path}")

    log(f"Model: {MODEL}")

    log(f"Batch size: {BATCH_SIZE}")

    log(f"Temperature: {TEMPERATURE}")

    log(f"Max output tokens: {MAX_TOKENS}")

    log()

    batch_responses = []

    all_records = []

    used_keys = set()

    model_types = {}

    try:

        log("Extracting text...")

        if extension == ".pdf":

            text, page_count = extract_pdf(input_path)

        else:

            text, page_count = extract_txt(input_path)

        extracted_path.write_text(text, encoding="utf-8")

        log(f"Extracted text saved to: {extracted_path}")

        pages = split_pages(text, page_count)

        batches = [

            pages[i:i + BATCH_SIZE]

            for i in range(0, len(pages), BATCH_SIZE)

        ]

        log(f"Total pages: {len(pages)}")

        log(f"Total batches: {len(batches)}")

        log()

        for batch_number, batch in enumerate(batches, start=1):

            page_numbers = [p["page"] for p in batch]

            log(f"[{batch_number}/{len(batches)}] Batch {batch_number}: pages {page_numbers}")

            batch_start = time.time()

            raw, parsed, prompt = process_batch(

                batch, batch_number, used_keys

            )

            inspect_fields(parsed, batch_number)

            records = parsed["records"]

            new_keys = collect_used_keys(records) - used_keys

            if new_keys:

                log("New approved keys: " + ", ".join(sorted(new_keys)))

            used_keys.update(collect_used_keys(records))

            batch_types = parsed.get("column_types", {})

            if isinstance(batch_types, dict):

                model_types.update(batch_types)

            all_records.extend(records)

            batch_responses.append({

                "batch_number": batch_number,

                "page_numbers": page_numbers,

                "prompt_kind": (

                    "recovery"

                    if prompt == build_recovery_prompt(batch, used_keys)

                    else ("discovery" if batch_number == 1 else "continuation")

                ),

                "prompt": prompt,

                "raw_response": raw,

                "parsed_response": parsed,

            })

            # Save progress after each successful batch.

            write_json(response_path, {

                "input_file": str(input_path),

                "batches": batch_responses,

            })

            log(f"Batch duration: {time.time() - batch_start:.2f} seconds")

            log(f"Records from batch: {len(records)}")

            log(f"Total records so far: {len(all_records)}")

            log()

        # Final safety check: never report success or write a blank workbook

        # when the model produced no usable facts across all batches.

        if not any(

            has_usable_extraction({"records": [record]})

            for record in all_records

        ):

            raise RuntimeError(

                "No usable facts were extracted from any batch. "

                "The Excel files were not written."

            )

        # Preserve the original model values and all model keys in raw Excel.

        raw_rows, raw_columns = flatten_raw(all_records)

        save_workbook(raw_rows, raw_columns, raw_path, convert=False)

        # Apply the fixed schema and type conversion only to final Excel.

        fixed_rows = flatten_fixed(all_records)

        for row_number, row in enumerate(fixed_rows, start=2):

            for key in row:

                if key not in FINAL_COLUMNS:

                    DROPPED_COLUMNS.append({

                        "row": row_number,

                        "column": key,

                    })

        dropped_names = sorted({

            item["column"] for item in DROPPED_COLUMNS

        })

        if dropped_names:

            log("Dropped columns from final Excel: " + ", ".join(dropped_names))

        save_workbook(

            fixed_rows,

            FINAL_COLUMNS,

            final_path,

            convert=True,

            model_types=model_types,

        )

        report_path = input_path.with_name(

            input_path.stem + "_hardening_report.json"

        )

        write_json(report_path, {

            "input_file": str(input_path),

            "model": MODEL,

            "invented_keys": INVENTED_KEYS,

            "type_warnings": TYPE_WARNINGS,

            "dropped_columns": DROPPED_COLUMNS,

            "record_count": len(all_records),

            "excel_row_count": len(fixed_rows),

            "elapsed_seconds": round(time.time() - started, 2),

        })

        elapsed = time.time() - started

        log()

        log("=" * 64)

        log("PIPELINE COMPLETED")

        log("=" * 64)

        log(f"End time: {datetime.now():%Y-%m-%d %H:%M:%S}")

        log(f"Total pages: {len(pages)}")

        log(f"Total batches: {len(batches)}")

        log(f"Total records: {len(all_records)}")

        log(f"Rows saved: {len(fixed_rows)}")

        log(f"Total duration: {elapsed:.2f} seconds")

        log()

        log(f"Extracted text: {extracted_path}")

        log(f"Response JSON: {response_path}")

        log(f"Raw Excel: {raw_path}")

        log(f"Final Excel: {final_path}")

        log(f"Hardening report: {report_path}")

        log()

        log(f"Unexpected key occurrences: {len(INVENTED_KEYS)}")

        log(f"Type conversion warnings: {len(TYPE_WARNINGS)}")

        log(f"Dropped columns: {len(DROPPED_COLUMNS)}")

        log("SUCCESS")

    except Exception as exc:

        # Keep successful batch responses for debugging if a later step fails.

        if batch_responses:

            write_json(response_path, {

                "input_file": str(input_path),

                "batches": batch_responses,

                "error": str(exc),

            })

        log()

        log("=" * 64)

        log("PIPELINE FAILED")

        log("=" * 64)

        log(f"ERROR: {exc}")

        raise

if __name__ == "__main__":

    try:

        main()

    except Exception:

        sys.exit(1)
