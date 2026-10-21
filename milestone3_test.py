import json
from pathlib import Path
import urllib.request

API_URL = "http://localhost:1234/v1/chat/completions"
MODEL = "qwen2.5-3b-instruct"

INPUT_FILE = Path("test_files/sample_invoice_extracted.txt")
OUTPUT_FILE = Path("test_files/sample_invoice_response.json")


COLUMN_LIST = """
HEADER COLUMNS

document_type:
The type of document, such as Invoice, Purchase Order, Receipt,
Delivery Note, Credit Note, or Debit Note.

document_number:
The document's own identifying number.

document_date:
The date the current document was issued or created.

seller_name:
The party issuing the document and supplying the goods or services.

seller_address:
The address belonging to the seller or issuer.

seller_tax_id:
The tax registration or tax ID belonging to the seller.

buyer_name:
The party purchasing or receiving the goods or services.

buyer_address:
The address belonging to the buyer or recipient.

buyer_tax_id:
The tax registration or tax ID belonging to the buyer.

reference_number:
A number referring to another document, order, transaction,
shipment, or external reference. Do not use the current document number.

due_date:
The date by which payment is due.

currency:
The currency used for monetary values.

subtotal_amount:
The document-level subtotal before final tax.

discount_amount:
A document-level discount amount explicitly printed in the document.
Never calculate it.

tax_amount:
A document-level tax amount explicitly printed in the document.
Never calculate it.

total_amount:
The final payable amount for the whole document.

remarks:
General document notes or remarks that do not fit another named field.


LINE ITEM COLUMNS

item_code:
The product, service, SKU, part, or item identifier for a line.

item_description:
The description or name of the product or service on a line.

quantity:
The quantity represented by the line item.

unit:
The unit of measurement associated with the quantity.

rate:
The per-unit price or rate for the line item.

line_amount:
The monetary amount associated with one specific line item.


EXTRA COLUMNS

extra_1 through extra_10:
Facts that do not fit any named column.
Use the same extra slot for the same concept if it appears again.
Do not use an extra column when a named column already fits.
"""


def main():
    document_text = INPUT_FILE.read_text(
        encoding="utf-8",
        errors="replace"
    )

    prompt = f"""
You are a document information extraction system.

Your job is to extract EVERY FACT in the document that matches
one of the columns below.

IMPORTANT:
1. Read the ENTIRE document before answering.
2. Do not stop after identifying the document number and date.
3. Extract ALL applicable header information.
4. Extract ALL product/service line items.
5. Map each value to the column whose DESCRIPTION best matches it.
6. The meaning of the value matters more than its position on the page.
7. Do not invent information.
8. Do not calculate missing values.
9. Do not rename columns.
10. Do not create new column names.
11. Do not put subtotal, tax, or grand-total rows into line_items.
12. If a value is not present, omit that field.
13. Preserve printed values when possible.
14. Return ONLY valid JSON.

COLUMN DEFINITIONS:
{COLUMN_LIST}

REQUIRED JSON STRUCTURE:

{{
  "records": [
    {{
      "page": 1,
      "header": {{
        "document_type": "...",
        "document_number": "...",
        "document_date": "...",
        "seller_name": "...",
        "seller_address": "...",
        "seller_tax_id": "...",
        "buyer_name": "...",
        "buyer_address": "...",
        "buyer_tax_id": "...",
        "reference_number": "...",
        "due_date": "...",
        "currency": "...",
        "subtotal_amount": "...",
        "discount_amount": "...",
        "tax_amount": "...",
        "total_amount": "...",
        "remarks": "...",
        "extra_1": "..."
      }},
      "line_items": [
        {{
          "item_code": "...",
          "item_description": "...",
          "quantity": "...",
          "unit": "...",
          "rate": "...",
          "line_amount": "..."
        }}
      ]
    }}
  ],
  "column_types": {{
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
    "remarks": "text",
    "item_code": "text",
    "item_description": "text",
    "quantity": "numeric",
    "unit": "text",
    "rate": "numeric",
    "line_amount": "numeric"
  }}
}}

DOCUMENT TO EXTRACT:

{document_text}
"""

    request_body = {
        "model": MODEL,
        "messages": [
            {
                "role": "system",
                "content": (
                    "You are a precise document extraction system. "
                    "Extract every applicable fact and return only JSON."
                )
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0,
    }

    request = urllib.request.Request(
        API_URL,
        data=json.dumps(request_body).encode("utf-8"),
        headers={
            "Content-Type": "application/json"
        },
        method="POST"
    )

    print("Sending document to LM Studio...")

    with urllib.request.urlopen(request, timeout=600) as response:
        response_data = json.loads(
            response.read().decode("utf-8")
        )

    content = response_data["choices"][0]["message"]["content"]

    parsed_json = json.loads(content)

    OUTPUT_FILE.write_text(
        json.dumps(
            parsed_json,
            indent=2,
            ensure_ascii=False
        ),
        encoding="utf-8"
    )

    print("\nSUCCESS")
    print(f"Response saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()