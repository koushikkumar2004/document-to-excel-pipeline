import json
import configparser
from pathlib import Path
from datetime import datetime
import time
import sys
import re

import requests
from openpyxl import Workbook


# ============================================================
# WINDOWS UTF-8 CONSOLE
# ============================================================

if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(
            encoding="utf-8",
            errors="replace"
        )
    except Exception:
        pass

if sys.stderr is not None:
    try:
        sys.stderr.reconfigure(
            encoding="utf-8",
            errors="replace"
        )
    except Exception:
        pass


# ============================================================
# CONFIGURATION
# ============================================================

CONFIG_FILE = "config.ini"

config = configparser.ConfigParser()

try:
    with open(
        CONFIG_FILE,
        "r",
        encoding="utf-8-sig"
    ) as file:
        config.read_file(file)

except FileNotFoundError:
    print(
        f"ERROR: Configuration file not found: "
        f"{CONFIG_FILE}"
    )
    sys.exit(1)

except configparser.Error as error:
    print(
        f"ERROR: Could not read config.ini: {error}"
    )
    sys.exit(1)


# ============================================================
# CONFIG HELPERS
# ============================================================

def get_required(section, key):
    if section not in config:
        raise RuntimeError(
            f"Missing [{section}] section in config.ini."
        )

    value = config[section].get(
        key,
        ""
    ).strip()

    if not value:
        raise RuntimeError(
            f"Missing required setting "
            f"[{section}] {key}."
        )

    return value


def get_optional(section, key, default):
    if section not in config:
        return default

    value = config[section].get(
        key,
        ""
    ).strip()

    if value == "":
        return default

    return value


def get_int(section, key, default):
    value = get_optional(
        section,
        key,
        str(default)
    )

    try:
        return int(value)
    except ValueError:
        raise RuntimeError(
            f"Invalid integer for "
            f"[{section}] {key}: {value}"
        )


def get_float(section, key, default):
    value = get_optional(
        section,
        key,
        str(default)
    )

    try:
        return float(value)
    except ValueError:
        raise RuntimeError(
            f"Invalid number for "
            f"[{section}] {key}: {value}"
        )


# ============================================================
# LM STUDIO SETTINGS
# ============================================================

BASE_URL = get_required(
    "lm_studio",
    "base_url"
)

MODEL = get_required(
    "lm_studio",
    "model"
)

TOKEN = get_optional(
    "lm_studio",
    "token",
    ""
)

TIMEOUT = get_int(
    "lm_studio",
    "timeout_seconds",
    600
)

TEMPERATURE = get_float(
    "lm_studio",
    "temperature",
    0.0
)

MAX_OUTPUT_TOKENS = get_int(
    "lm_studio",
    "max_output_tokens",
    4096
)


# ============================================================
# PIPELINE SETTINGS
# ============================================================

BATCH_SIZE = get_int(
    "pipeline",
    "batch_size",
    1
)

SPARSE_PAGE_THRESHOLD = get_int(
    "pipeline",
    "sparse_page_threshold",
    20
)

EXTRA_SLOT_COUNT = get_int(
    "pipeline",
    "extra_slot_count",
    10
)

PDFTOTEXT_PATH = get_optional(
    "pipeline",
    "pdftotext_path",
    "bin/pdftotext.exe"
)


# ============================================================
# VALIDATE CONFIGURATION
# ============================================================

if BATCH_SIZE < 1:
    raise RuntimeError(
        "batch_size must be at least 1."
    )

if TIMEOUT < 1:
    raise RuntimeError(
        "timeout_seconds must be at least 1."
    )

if MAX_OUTPUT_TOKENS < 1:
    raise RuntimeError(
        "max_output_tokens must be at least 1."
    )

if EXTRA_SLOT_COUNT < 0:
    raise RuntimeError(
        "extra_slot_count cannot be negative."
    )


API_URL = (
    BASE_URL.rstrip("/")
    + "/chat/completions"
)


# ============================================================
# FIXED COLUMN LIST
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
    f"extra_{i}"
    for i in range(
        1,
        EXTRA_SLOT_COUNT + 1
    )
]


ALLOWED_COLUMNS = (
    HEADER_COLUMNS
    + LINE_ITEM_COLUMNS
    + EXTRA_COLUMNS
)


ALLOWED_COLUMN_SET = set(
    ALLOWED_COLUMNS
)


# ============================================================
# COLUMN TYPES
# ============================================================

COLUMN_TYPES = {

    "document_type": "text",
    "document_number": "text",
    "document_date": "date",

    "seller_name": "text",
    "seller_address": "text",
    "seller_tax_id": "text",

    "buyer_name": "text",
    "buyer_address": "text",
    "buyer_tax_id": "text",

    "reference_number": "text",
    "due_date": "date",

    "currency": "text",

    "subtotal_amount": "numeric",
    "discount_amount": "numeric",
    "tax_amount": "numeric",
    "total_amount": "numeric",

    "item_code": "text",
    "item_description": "text",
    "quantity": "numeric",
    "unit": "text",
    "rate": "numeric",
    "line_amount": "numeric",

    "remarks": "text",
}


# ============================================================
# LOGGING HELPERS
# ============================================================

INVENTED_KEYS = []
DROPPED_COLUMNS = []
TYPE_WARNINGS = []
NEW_KEYS_BY_BATCH = []


def log_invented_key(
    key,
    batch_number,
    location
):

    entry = {
        "key": key,
        "batch": batch_number,
        "location": location
    }

    INVENTED_KEYS.append(
        entry
    )

    print(
        f"WARNING: Invented/new key "
        f"'{key}' found in batch "
        f"{batch_number} at {location}."
    )


def log_dropped_column(
    column,
    row_number
):

    entry = {
        "column": column,
        "row": row_number
    }

    DROPPED_COLUMNS.append(
        entry
    )

    print(
        f"WARNING: Dropped column "
        f"'{column}' from Excel row "
        f"{row_number}."
    )


def log_type_warning(
    column,
    row_number,
    value,
    expected_type
):

    entry = {
        "column": column,
        "row": row_number,
        "value": value,
        "expected_type": expected_type
    }

    TYPE_WARNINGS.append(
        entry
    )

    print(
        f"WARNING: Type conversion failed "
        f"for column '{column}', "
        f"row {row_number}, "
        f"value '{value}', "
        f"expected {expected_type}."
    )


# ============================================================
# FIND KEYS RECURSIVELY
# ============================================================

def validate_keys(
    obj,
    batch_number,
    location="root"
):

    """
    Validate every dictionary key against
    the fixed schema.

    IMPORTANT:
    This function only detects and logs.
    It does NOT rename or move anything.
    """

    if isinstance(obj, dict):

        for key, value in obj.items():

            # These are structural JSON keys,
            # not Excel columns.
            structural_keys = {
                "records",
                "page",
                "header",
                "line_items",
                "column_types"
            }

            if (
                key not in ALLOWED_COLUMN_SET
                and key not in structural_keys
            ):

                log_invented_key(
                    key,
                    batch_number,
                    location
                )

            validate_keys(
                value,
                batch_number,
                f"{location}.{key}"
            )

    elif isinstance(obj, list):

        for index, item in enumerate(obj):

            validate_keys(
                item,
                batch_number,
                f"{location}[{index}]"
            )


# ============================================================
# EXTRACT KEYS USED BY BATCH
# ============================================================

def collect_used_column_keys(obj):

    found = set()

    def walk(value):

        if isinstance(value, dict):

            for key, child in value.items():

                if key in ALLOWED_COLUMN_SET:
                    found.add(key)

                walk(child)

        elif isinstance(value, list):

            for item in value:
                walk(item)

    walk(obj)

    return found


# ============================================================
# LOAD PAGES
# ============================================================

def load_pages(file_path):

    file_path = Path(file_path)

    if not file_path.exists():
        raise RuntimeError(
            f"Input extracted text file not found: "
            f"{file_path}"
        )

    text = file_path.read_text(
        encoding="utf-8-sig",
        errors="replace"
    )

    raw_pages = text.split(
        "--- Page "
    )

    pages = []

    for block in raw_pages:

        block = block.strip()

        if not block:
            continue

        marker_end = block.find(
            "---"
        )

        if marker_end == -1:
            continue

        page_number_text = (
            block[:marker_end]
            .strip()
        )

        try:
            page_number = int(
                page_number_text
            )

        except ValueError:
            continue

        page_text = (
            block[marker_end + 3:]
            .strip()
        )

        non_whitespace = len(
            "".join(
                page_text.split()
            )
        )

        if (
            non_whitespace
            < SPARSE_PAGE_THRESHOLD
        ):

            print(
                f"WARNING: Page {page_number} "
                f"contains very little text "
                f"({non_whitespace} characters)."
            )

        pages.append(
            {
                "page": page_number,
                "text": page_text
            }
        )

    if not pages:
        raise RuntimeError(
            "No pages found in extracted text."
        )

    return pages


# ============================================================
# CREATE BATCHES
# ============================================================

def create_batches(
    pages,
    batch_size
):

    batches = []

    for i in range(
        0,
        len(pages),
        batch_size
    ):

        batches.append(
            pages[
                i:i + batch_size
            ]
        )

    return batches


# ============================================================
# REMOVE THINKING / CODE FENCES
# ============================================================

def clean_llm_content(content):

    if not isinstance(
        content,
        str
    ):

        raise RuntimeError(
            "LM Studio content is not text."
        )

    content = content.strip()

    if not content:
        raise RuntimeError(
            "LM Studio returned empty content."
        )

    # --------------------------------------------------------
    # Remove common thinking/reasoning blocks
    # --------------------------------------------------------

    content = re.sub(
        r"<think>.*?</think>",
        "",
        content,
        flags=re.DOTALL |
        re.IGNORECASE
    ).strip()

    # --------------------------------------------------------
    # Remove markdown code fences
    # --------------------------------------------------------

    if content.startswith("```"):

        lines = content.splitlines()

        if (
            lines
            and lines[0]
            .strip()
            .startswith("```")
        ):
            lines = lines[1:]

        if (
            lines
            and lines[-1]
            .strip()
            == "```"
        ):
            lines = lines[:-1]

        content = "\n".join(
            lines
        ).strip()

    return content


# ============================================================
# PARSE JSON
# ============================================================

def parse_json_content(
    content
):

    cleaned = clean_llm_content(
        content
    )

    try:

        return json.loads(
            cleaned
        )

    except json.JSONDecodeError:

        start = cleaned.find(
            "{"
        )

        end = cleaned.rfind(
            "}"
        )

        if (
            start == -1
            or end == -1
            or end <= start
        ):

            raise RuntimeError(
                "No JSON object found "
                "in LM Studio response."
            )

        candidate = cleaned[
            start:end + 1
        ]

        try:

            return json.loads(
                candidate
            )

        except json.JSONDecodeError as error:

            raise RuntimeError(
                f"Invalid JSON: {error}"
            )


# ============================================================
# VALIDATE LLM API RESPONSE
# ============================================================

def parse_llm_response(
    response
):

    if not isinstance(
        response,
        dict
    ):

        raise RuntimeError(
            "LM Studio API response "
            "is not a JSON object."
        )

    choices = response.get(
        "choices"
    )

    if not choices:

        raise RuntimeError(
            "LM Studio response is missing "
            "'choices'."
        )

    if not isinstance(
        choices,
        list
    ):

        raise RuntimeError(
            "LM Studio 'choices' is not a list."
        )

    first_choice = choices[0]

    if not isinstance(
        first_choice,
        dict
    ):

        raise RuntimeError(
            "LM Studio choices[0] is invalid."
        )

    finish_reason = first_choice.get(
        "finish_reason"
    )

    message = first_choice.get(
        "message"
    )

    if not isinstance(
        message,
        dict
    ):

        raise RuntimeError(
            "LM Studio response is missing "
            "'message'."
        )

    content = message.get(
        "content"
    )

    if content is None:

        raise RuntimeError(
            "LM Studio response is missing "
            "'content'."
        )

    if not str(content).strip():

        reason = finish_reason or "unknown"

        raise RuntimeError(
            "LM Studio returned empty content. "
            f"Finish reason: {reason}"
        )

    parsed = parse_json_content(
        content
    )

    if not isinstance(
        parsed,
        dict
    ):

        raise RuntimeError(
            "Parsed response is not a JSON object."
        )

    if "records" not in parsed:

        raise RuntimeError(
            "Parsed JSON does not contain "
            "'records'."
        )

    if not isinstance(
        parsed["records"],
        list
    ):

        raise RuntimeError(
            "'records' must be a list."
        )

    return parsed


# ============================================================
# BUILD PROMPT
# ============================================================

def build_prompt(
    batch,
    previously_used_keys
):

    pages_text = []

    for page in batch:

        pages_text.append(
            f"--- Page {page['page']} ---\n"
            f"{page['text']}"
        )

    combined_text = "\n\n".join(
        pages_text
    )

    continuation_info = ""

    if previously_used_keys:

        continuation_info = (
            "\nPreviously used column keys:\n"
            + "\n".join(
                sorted(
                    previously_used_keys
                )
            )
            + "\nReuse these keys for the "
            "same concepts. Do not change "
            "their meaning.\n"
        )

    prompt = f"""
You are a document information extraction system.

Extract facts from the supplied document pages.

Map facts ONLY to these fixed columns:

{chr(10).join(ALLOWED_COLUMNS)}

The mapping must be decided by the LLM from the
meaning of the document.

Python code must NOT rename, move, calculate,
guess, or repair values.

If a value is not present, use an empty string
or null.

For line-item documents:

- Create one line_items object per actual item.
- Do not treat subtotal, tax, total, or grand-total
  rows as line items.
- A fee/charge row that belongs to the item above
  should remain associated with that item.
- UNIT belongs in unit.
- ITEM DESCRIPTION belongs in item_description.
- QUANTITY belongs in quantity.
- Do not put UNIT into item_code.
- Do not put ITEM DESCRIPTION into item_code.

Do not calculate rate or line_amount when they
are not present.

Each field should contain one atomic fact.

Strip printed labels from values.

Do not record placeholders such as:
N/A
-
see below

Never duplicate the same printed occurrence.

{continuation_info}

Return ONLY valid JSON.

Required structure:

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

Document:

{combined_text}
"""

    return prompt


# ============================================================
# SEND ONE LM STUDIO REQUEST
# ============================================================

def send_request(
    prompt
):

    headers = {
        "Content-Type": "application/json"
    }

    if TOKEN:

        headers[
            "Authorization"
        ] = f"Bearer {TOKEN}"

    payload = {

        "model": MODEL,

        "messages": [
            {
                "role": "system",
                "content": (
                    "You extract document facts "
                    "into strict JSON."
                )
            },
            {
                "role": "user",
                "content": prompt
            }
        ],

        "temperature": TEMPERATURE,

        "max_tokens": MAX_OUTPUT_TOKENS
    }

    try:

        response = requests.post(
            API_URL,
            headers=headers,
            json=payload,
            timeout=TIMEOUT
        )

    except requests.exceptions.Timeout:

        raise RuntimeError(
            f"LM Studio request timed out "
            f"after {TIMEOUT} seconds."
        )

    except requests.exceptions.ConnectionError:

        raise RuntimeError(
            "LM Studio is unreachable. "
            "Make sure the local server is running."
        )

    except requests.exceptions.RequestException as error:

        raise RuntimeError(
            f"LM Studio request failed: "
            f"{error}"
        )

    if response.status_code == 401:

        raise RuntimeError(
            "LM Studio authentication failed "
            "(HTTP 401). Check the token."
        )

    if response.status_code != 200:

        raise RuntimeError(
            f"LM Studio returned HTTP "
            f"{response.status_code}:\n"
            f"{response.text}"
        )

    try:

        return response.json()

    except ValueError:

        raise RuntimeError(
            "LM Studio HTTP response is not "
            "valid JSON."
        )


# ============================================================
# CALL LM STUDIO WITH ONE RETRY
# ============================================================

def call_lm_studio(
    batch,
    previously_used_keys,
    batch_number
):

    prompt = build_prompt(
        batch,
        previously_used_keys
    )

    last_error = None
    last_raw_content = None

    for attempt in range(
        1,
        3
    ):

        print(
            f"  LM Studio attempt "
            f"{attempt}/2..."
        )

        try:

            raw_response = send_request(
                prompt
            )

            # ------------------------------------------------
            # Preserve raw API response
            # ------------------------------------------------

            choices = raw_response.get(
                "choices",
                []
            )

            if choices:

                message = choices[0].get(
                    "message",
                    {}
                )

                if isinstance(
                    message,
                    dict
                ):

                    last_raw_content = (
                        message.get(
                            "content"
                        )
                    )

            parsed = parse_llm_response(
                raw_response
            )

            return (
                raw_response,
                parsed,
                prompt
            )

        except Exception as error:

            last_error = error

            print(
                f"  Attempt {attempt} failed: "
                f"{error}"
            )

            if attempt == 1:

                print(
                    "  Retrying once..."
                )

    raw_preview = (
        str(last_raw_content)
        if last_raw_content
        else "<no raw content>"
    )

    raise RuntimeError(
        "LM Studio response remained invalid "
        "after retry.\n"
        f"Reason: {last_error}\n"
        f"Raw content:\n{raw_preview}"
    )


# ============================================================
# CONVERT VALUE TO NUMBER
# ============================================================

def convert_number(
    value,
    column,
    row_number
):

    if value is None:
        return None

    if isinstance(
        value,
        bool
    ):
        return value

    if isinstance(
        value,
        (int, float)
    ):
        return value

    original = str(
        value
    ).strip()

    if not original:
        return ""

    cleaned = original

    # Currency prefixes
    cleaned = re.sub(
        r"^(INR|Rs\.?|₹|\$|€|£)\s*",
        "",
        cleaned,
        flags=re.IGNORECASE
    )

    # Percent sign
    cleaned = cleaned.rstrip(
        "%"
    ).strip()

    # Thousands separators
    cleaned = cleaned.replace(
        ",",
        ""
    )

    try:

        number = float(
            cleaned
        )

        if number.is_integer():
            return int(number)

        return number

    except ValueError:

        log_type_warning(
            column,
            row_number,
            value,
            "numeric"
        )

        return value


# ============================================================
# CONVERT VALUE TO DATE
# ============================================================

def convert_date(
    value,
    column,
    row_number
):

    if value is None:
        return None

    if hasattr(
        value,
        "date"
    ):

        try:
            return value.date()
        except Exception:
            pass

    if not isinstance(
        value,
        str
    ):

        log_type_warning(
            column,
            row_number,
            value,
            "date"
        )

        return value

    text = value.strip()

    if not text:
        return ""

    # ISO datetime
    iso_match = re.match(
        r"^(\d{4}-\d{2}-\d{2})",
        text
    )

    if iso_match:

        try:

            from datetime import date

            return date.fromisoformat(
                iso_match.group(1)
            )

        except ValueError:
            pass

    # Day-first formats
    date_patterns = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%d/%m/%y",
        "%d-%m-%y",
        "%d %B %Y",
        "%d %b %Y"
    ]

    from datetime import datetime as dt

    for pattern in date_patterns:

        try:

            return dt.strptime(
                text,
                pattern
            ).date()

        except ValueError:
            continue

    log_type_warning(
        column,
        row_number,
        value,
        "date"
    )

    return value


# ============================================================
# TYPE CONVERSION
# ============================================================

def convert_value(
    column,
    value,
    row_number,
    model_column_types
):

    if column in COLUMN_TYPES:

        expected_type = (
            COLUMN_TYPES[column]
        )

    else:

        expected_type = (
            model_column_types.get(
                column,
                "text"
            )
        )

    if expected_type == "numeric":

        return convert_number(
            value,
            column,
            row_number
        )

    if expected_type == "date":

        return convert_date(
            value,
            column,
            row_number
        )

    return value


# ============================================================
# FLATTEN RECORDS
# ============================================================

def flatten_records(
    records,
    model_column_types
):

    rows = []

    for record in records:

        page = record.get(
            "page"
        )

        header = record.get(
            "header",
            {}
        )

        line_items = record.get(
            "line_items",
            []
        )

        if not isinstance(
            header,
            dict
        ):
            header = {}

        if not isinstance(
            line_items,
            list
        ):
            line_items = []

        # ----------------------------------------------------
        # Record with line items
        # ----------------------------------------------------

        if line_items:

            for item in line_items:

                row = {}

                # Header fields
                for column in HEADER_COLUMNS:

                    row[column] = header.get(
                        column,
                        ""
                    )

                # Line item fields
                if isinstance(
                    item,
                    dict
                ):

                    for column in LINE_ITEM_COLUMNS:

                        row[column] = item.get(
                            column,
                            ""
                        )

                # Extra slots
                for column in EXTRA_COLUMNS:

                    row[column] = header.get(
                        column,
                        ""
                    )

                row[
                    "source_page"
                ] = page

                rows.append(
                    row
                )

        # ----------------------------------------------------
        # Record without line items
        # ----------------------------------------------------

        else:

            row = {}

            for column in HEADER_COLUMNS:

                row[column] = header.get(
                    column,
                    ""
                )

            for column in LINE_ITEM_COLUMNS:

                row[column] = ""

            for column in EXTRA_COLUMNS:

                row[column] = header.get(
                    column,
                    ""
                )

            row[
                "source_page"
            ] = page

            rows.append(
                row
            )

    return rows


# ============================================================
# SAVE RAW EXCEL
# ============================================================

def save_raw_excel(
    records,
    output_path,
    model_column_types
):

    rows = flatten_records(
        records,
        model_column_types
    )

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Raw Data"

    # --------------------------------------------------------
    # Determine every key actually used
    # --------------------------------------------------------

    all_keys = set()

    def collect(value):

        if isinstance(
            value,
            dict
        ):

            for key, child in value.items():

                if key in ALLOWED_COLUMN_SET:

                    all_keys.add(
                        key
                    )

                collect(child)

        elif isinstance(
            value,
            list
        ):

            for item in value:
                collect(item)

    collect(records)

    # Fixed columns first
    raw_columns = list(
        ALLOWED_COLUMNS
    )

    # Then any unexpected columns
    unexpected = sorted(
        all_keys
        - ALLOWED_COLUMN_SET
    )

    # The unexpected keys are normally
    # structural/invented fields and have
    # already been logged.
    raw_columns.extend(
        unexpected
    )

    raw_columns.append(
        "source_page"
    )

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    for column_number, column in enumerate(
        raw_columns,
        start=1
    ):

        sheet.cell(
            row=1,
            column=column_number,
            value=column
        )

    # --------------------------------------------------------
    # Rows
    # --------------------------------------------------------

    for row_number, row in enumerate(
        rows,
        start=2
    ):

        for column_number, column in enumerate(
            raw_columns,
            start=1
        ):

            if column == "source_page":

                value = row.get(
                    "source_page",
                    ""
                )

            else:

                if column in ALLOWED_COLUMN_SET:

                    value = row.get(
                        column,
                        ""
                    )

                else:

                    value = ""

                    log_dropped_column(
                        column,
                        row_number
                    )

            sheet.cell(
                row=row_number,
                column=column_number,
                value=value
            )

    # --------------------------------------------------------
    # Apply type conversion to the fixed columns
    # --------------------------------------------------------

    for row_number in range(
        2,
        sheet.max_row + 1
    ):

        for column_number, column in enumerate(
            raw_columns,
            start=1
        ):

            if column not in ALLOWED_COLUMN_SET:
                continue

            cell = sheet.cell(
                row=row_number,
                column=column_number
            )

            converted = convert_value(
                column,
                cell.value,
                row_number,
                model_column_types
            )

            cell.value = converted

    workbook.save(
        output_path
    )

    return len(rows)


# ============================================================
# SAVE DEBUG REPORT
# ============================================================

def save_debug_report(
    output_dir,
    input_file,
    start_timestamp,
    end_timestamp,
    duration,
    page_count,
    batch_count,
    record_count
):

    report = {

        "input_file": str(
            input_file
        ),

        "start_time": start_timestamp,

        "end_time": end_timestamp,

        "duration_seconds": round(
            duration,
            2
        ),

        "page_count": page_count,

        "batch_count": batch_count,

        "record_count": record_count,

        "invented_keys": INVENTED_KEYS,

        "type_conversion_warnings": TYPE_WARNINGS,

        "dropped_columns": DROPPED_COLUMNS,

        "allowed_columns": ALLOWED_COLUMNS
    }

    report_path = (
        output_dir
        / "sample_invoice12_hardening_report.json"
    )

    report_path.write_text(
        json.dumps(
            report,
            indent=2,
            ensure_ascii=False,
            default=str
        ),
        encoding="utf-8"
    )

    return report_path


# ============================================================
# MAIN
# ============================================================

def main():

    start_time = time.time()

    start_timestamp = (
        datetime.now()
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    print("=" * 60)
    print(
        "DOCUMENT-TO-EXCEL PIPELINE"
    )
    print("=" * 60)

    print(
        f"Start time: {start_timestamp}"
    )

    print(
        f"Model: {MODEL}"
    )

    print(
        f"Batch size: {BATCH_SIZE}"
    )

    print(
        f"Temperature: {TEMPERATURE}"
    )

    print(
        f"Max output tokens: "
        f"{MAX_OUTPUT_TOKENS}"
    )

    print(
        f"Timeout: {TIMEOUT} seconds"
    )

    print()

    # --------------------------------------------------------
    # Input/output paths
    # --------------------------------------------------------

    input_file = Path(
        "test_files/"
        "sample_invoice12_extracted.txt"
    )

    output_dir = Path(
        "test_files"
    )

    output_dir.mkdir(
        exist_ok=True
    )

    # --------------------------------------------------------
    # Load pages
    # --------------------------------------------------------

    pages = load_pages(
        input_file
    )

    print(
        f"Total pages: {len(pages)}"
    )

    # --------------------------------------------------------
    # Create batches
    # --------------------------------------------------------

    batches = create_batches(
        pages,
        BATCH_SIZE
    )

    print(
        f"Total batches: {len(batches)}"
    )

    print()

    # --------------------------------------------------------
    # Combined raw responses
    # --------------------------------------------------------

    raw_batch_responses = []

    all_records = []

    previously_used_keys = set()

    # --------------------------------------------------------
    # Process each batch
    # --------------------------------------------------------

    for batch_number, batch in enumerate(
        batches,
        start=1
    ):

        page_numbers = [
            page["page"]
            for page in batch
        ]

        print(
            f"[{batch_number}/{len(batches)}] "
            f"Batch {batch_number}: "
            f"pages {page_numbers}"
        )

        batch_start = time.time()

        try:

            (
                raw_response,
                parsed_result,
                prompt
            ) = call_lm_studio(
                batch,
                previously_used_keys,
                batch_number
            )

            # ------------------------------------------------
            # Validate keys
            # ------------------------------------------------

            validate_keys(
                parsed_result,
                batch_number
            )

            current_keys = (
                collect_used_column_keys(
                    parsed_result
                )
            )

            new_keys = (
                current_keys
                - previously_used_keys
            )

            if new_keys:

                print(
                    "New approved keys in this batch: "
                    + ", ".join(
                        sorted(new_keys)
                    )
                )

                NEW_KEYS_BY_BATCH.append(
                    {
                        "batch": batch_number,
                        "keys": sorted(
                            new_keys
                        )
                    }
                )

            previously_used_keys.update(
                current_keys
            )

            # ------------------------------------------------
            # Save parsed batch result
            # ------------------------------------------------

            parsed_output = (
                output_dir
                / (
                    "sample_invoice12_"
                    f"batch{batch_number}_"
                    "response.json"
                )
            )

            parsed_output.write_text(
                json.dumps(
                    parsed_result,
                    indent=2,
                    ensure_ascii=False
                ),
                encoding="utf-8"
            )

            # ------------------------------------------------
            # Preserve raw API response
            # ------------------------------------------------

            raw_batch_responses.append(
                {
                    "batch_number": batch_number,
                    "page_numbers": page_numbers,
                    "prompt": prompt,
                    "response": raw_response
                }
            )

            # ------------------------------------------------
            # Collect records
            # ------------------------------------------------

            records = parsed_result.get(
                "records",
                []
            )

            all_records.extend(
                records
            )

            batch_duration = (
                time.time()
                - batch_start
            )

            print(
                f"Batch {batch_number} completed "
                f"in {batch_duration:.2f} seconds."
            )

            print(
                f"Records from batch: "
                f"{len(records)}"
            )

            print(
                f"Total records so far: "
                f"{len(all_records)}"
            )

        except Exception as error:

            batch_duration = (
                time.time()
                - batch_start
            )

            print()
            print(
                f"ERROR in batch "
                f"{batch_number}: {error}"
            )

            print(
                f"Batch runtime before failure: "
                f"{batch_duration:.2f} seconds"
            )

            raise

        print()

    # ========================================================
    # SAVE COMBINED RESPONSE
    # ========================================================

    combined_response = {

        "records": all_records,

        "column_types": {},

        "batches": raw_batch_responses
    }

    combined_output = (
        output_dir
        / "sample_invoice12_combined_response.json"
    )

    combined_output.write_text(
        json.dumps(
            combined_response,
            indent=2,
            ensure_ascii=False
        ),
        encoding="utf-8"
    )

    # ========================================================
    # RAW EXCEL
    # ========================================================

    model_column_types = {}

    for record in raw_batch_responses:

        response = record.get(
            "response",
            {}
        )

        choices = response.get(
            "choices",
            []
        )

        if not choices:
            continue

        message = choices[0].get(
            "message",
            {}
        )

        content = message.get(
            "content",
            ""
        )

        try:

            parsed = parse_json_content(
                content
            )

            types = parsed.get(
                "column_types",
                {}
            )

            if isinstance(
                types,
                dict
            ):

                model_column_types.update(
                    types
                )

        except Exception:
            pass

    raw_excel_output = (
        output_dir
        / "sample_invoice12_raw.xlsx"
    )

    row_count = save_raw_excel(
        all_records,
        raw_excel_output,
        model_column_types
    )

    # ========================================================
    # FINISH
    # ========================================================

    end_time = time.time()

    end_timestamp = (
        datetime.now()
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    duration = (
        end_time
        - start_time
    )

    # --------------------------------------------------------
    # Debug report
    # --------------------------------------------------------

    report_path = save_debug_report(
        output_dir,
        input_file,
        start_timestamp,
        end_timestamp,
        duration,
        len(pages),
        len(batches),
        len(all_records)
    )

    # ========================================================
    # FINAL LOG
    # ========================================================

    print("=" * 60)
    print(
        "PIPELINE COMPLETED"
    )
    print("=" * 60)

    print(
        f"End time: {end_timestamp}"
    )

    print(
        f"Total pages: {len(pages)}"
    )

    print(
        f"Total batches: {len(batches)}"
    )

    print(
        f"Total records: {len(all_records)}"
    )

    print(
        f"Rows saved to raw Excel: "
        f"{row_count}"
    )

    print(
        f"Total duration: "
        f"{duration:.2f} seconds"
    )

    print()

    print(
        f"Combined response: "
        f"{combined_output}"
    )

    print(
        f"Raw Excel: "
        f"{raw_excel_output}"
    )

    print(
        f"Hardening report: "
        f"{report_path}"
    )

    print()

    print(
        f"Invented/new keys detected: "
        f"{len(INVENTED_KEYS)}"
    )

    print(
        f"Type conversion warnings: "
        f"{len(TYPE_WARNINGS)}"
    )

    print(
        f"Dropped columns: "
        f"{len(DROPPED_COLUMNS)}"
    )

    print()

    print("SUCCESS")


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print()
        print(
            "ERROR: Pipeline interrupted by user."
        )

        sys.exit(130)

    except Exception as error:

        print()
        print(
            f"ERROR: {error}"
        )

        sys.exit(1)