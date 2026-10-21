import json
from pathlib import Path
from datetime import datetime

from openpyxl import Workbook


INPUT_FILE = Path(
    "test_files/sample_invoice12_combined_response.json"
)

OUTPUT_FILE = Path(
    "test_files/sample_invoice12_raw.xlsx"
)


# Fixed column order.
# These are the semantic columns decided by the column list.
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
    "extra_1",
    "extra_2",
    "extra_3",
    "extra_4",
    "extra_5",
    "extra_6",
    "extra_7",
    "extra_8",
    "extra_9",
    "extra_10",
]

ALL_COLUMNS = (
    HEADER_COLUMNS
    + LINE_ITEM_COLUMNS
    + EXTRA_COLUMNS
    + ["source_page"]
)


def convert_date(value):
    """
    Convert common date strings into real Excel dates.

    If conversion fails, preserve the original value.
    """

    if value is None or value == "":
        return value

    if isinstance(value, datetime):
        return value

    if not isinstance(value, str):
        return value

    formats = [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d %B %Y",
        "%d %b %Y",
    ]

    for date_format in formats:
        try:
            return datetime.strptime(
                value.strip(),
                date_format
            )
        except ValueError:
            pass

    print(
        f"WARNING: Could not convert date: {value}"
    )

    return value


def convert_numeric(value):
    """
    Convert numeric-looking values into real Excel numbers.

    If conversion fails, preserve the original value.
    """

    if value is None or value == "":
        return value

    if isinstance(value, (int, float)):
        return value

    if not isinstance(value, str):
        return value

    cleaned = value.strip()

    # Remove common currency symbols/prefixes.
    for prefix in [
        "₹",
        "$",
        "€",
        "£",
        "INR",
        "USD",
        "EUR",
        "GBP",
    ]:
        if cleaned.upper().startswith(prefix):
            cleaned = cleaned[len(prefix):].strip()

    # Remove thousands separators.
    cleaned = cleaned.replace(",", "")

    # Remove trailing percentage symbol.
    if cleaned.endswith("%"):
        cleaned = cleaned[:-1].strip()

    try:
        number = float(cleaned)

        if number.is_integer():
            return int(number)

        return number

    except ValueError:

        print(
            f"WARNING: Could not convert numeric value: {value}"
        )

        return value


def convert_value(column_name, value):

    if value is None:
        return ""

    if column_name in [
        "document_date",
        "due_date",
    ]:
        return convert_date(value)

    if column_name in [
        "subtotal_amount",
        "discount_amount",
        "tax_amount",
        "total_amount",
        "quantity",
        "rate",
        "line_amount",
    ]:
        return convert_numeric(value)

    return value


def create_row(record, line_item=None):

    row = {}

    header = record.get(
        "header",
        {}
    )

    # Header values.
    for column in HEADER_COLUMNS:

        value = header.get(
            column,
            ""
        )

        row[column] = convert_value(
            column,
            value
        )

    # Line-item values.
    if line_item is not None:

        for column in LINE_ITEM_COLUMNS:

            value = line_item.get(
                column,
                ""
            )

            row[column] = convert_value(
                column,
                value
            )

    else:

        for column in LINE_ITEM_COLUMNS:
            row[column] = ""

    # Extra slots.
    for column in EXTRA_COLUMNS:

        value = header.get(
            column,
            ""
        )

        row[column] = value

    # source_page is bookkeeping information.
    row["source_page"] = record.get(
        "page",
        ""
    )

    return row


def flatten_records(data):

    rows = []

    records = data.get(
        "records",
        []
    )

    for record in records:

        line_items = record.get(
            "line_items",
            []
        )

        # If line items exist:
        # create one Excel row per line item.
        if line_items:

            for line_item in line_items:

                row = create_row(
                    record,
                    line_item
                )

                rows.append(row)

        else:

            # No line items:
            # still create one row for the record.
            row = create_row(
                record
            )

            rows.append(row)

    return rows


def write_excel(rows):

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Extracted Data"

    # Write headers.
    for column_number, column_name in enumerate(
        ALL_COLUMNS,
        start=1
    ):

        worksheet.cell(
            row=1,
            column=column_number,
            value=column_name
        )

    # Write data.
    for row_number, row_data in enumerate(
        rows,
        start=2
    ):

        for column_number, column_name in enumerate(
            ALL_COLUMNS,
            start=1
        ):

            worksheet.cell(
                row=row_number,
                column=column_number,
                value=row_data.get(
                    column_name,
                    ""
                )
            )

    # Apply date format to date columns.
    for row in worksheet.iter_rows(
        min_row=2
    ):

        for cell in row:

            if cell.column in [
                ALL_COLUMNS.index("document_date") + 1,
                ALL_COLUMNS.index("due_date") + 1,
            ]:

                if isinstance(
                    cell.value,
                    datetime
                ):

                    cell.number_format = "dd/mm/yyyy"

    workbook.save(
        OUTPUT_FILE
    )


def main():

    print("Loading LLM JSON...")

    data = json.loads(
        INPUT_FILE.read_text(
            encoding="utf-8"
        )
    )

    print(
        f"Records found: "
        f"{len(data.get('records', []))}"
    )

    rows = flatten_records(
        data
    )

    print(
        f"Excel rows created: "
        f"{len(rows)}"
    )

    write_excel(
        rows
    )

    print("\nSUCCESS")
    print(
        f"Excel saved to: "
        f"{OUTPUT_FILE}"
    )


if __name__ == "__main__":
    main()