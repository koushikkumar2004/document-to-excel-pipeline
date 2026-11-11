# Document-to-Excel Pipeline

Converts text-based PDF invoices into structured Excel spreadsheets using a
locally hosted LLM (Qwen 2.5) via [LM Studio](https://lmstudio.ai/).

---

## Overview

```
PDF / TXT  →  pdftotext  →  LLM (LM Studio)  →  JSON  →  Excel (.xlsx)
```

The pipeline:
1. Extracts text from the PDF using the bundled `pdftotext` (Poppler).
2. Splits the text into page batches and sends each batch to the LLM.
3. Validates the structured JSON response and retries up to 3 times per batch.
4. Maps all fields to a fixed Excel schema (17 header + 6 line-item columns).
5. Writes the Excel file and a diagnostic JSON report alongside it.

---

## Prerequisites

| Requirement | Version tested |
|-------------|---------------|
| Python | 3.13.2 |
| [LM Studio](https://lmstudio.ai/) | ≥ 0.3 |
| Model loaded in LM Studio | `qwen2.5-3b-instruct` |
| Windows | 10 / 11 (Poppler binaries bundled in `bin/`) |

---

## Installation

```powershell
# 1. Clone / copy the project folder
cd document_excel_pipeline

# 2. Install Python dependencies
pip install -r requirements.txt
```

> [!NOTE]
> `pdftotext.exe` (Poppler for Windows) is already bundled in the `bin/`
> directory — no separate Poppler installation is required.

---

## LM Studio Setup

1. Download and install [LM Studio](https://lmstudio.ai/).
2. Inside LM Studio, download **Qwen 2.5 3B Instruct** (search: `qwen2.5-3b-instruct`).
3. Start the **Local Server** tab — the default endpoint is `http://localhost:1234/v1`.
4. Load the Qwen model so the server shows **"Model loaded"**.

> [!IMPORTANT]
> The pipeline will time out after 180 seconds per LLM call. If your machine
> is slow, increase `timeout_seconds` in `config.ini`.

---

## Configuration

All settings live in [`config.ini`](config.ini).  
A safe template with comments is provided as [`config.example.ini`](config.example.ini).

```ini
[lm_studio]
base_url        = http://localhost:1234/v1   # LM Studio server URL
api_token       = lm-studio                  # Any non-empty string
model           = qwen2.5-3b-instruct        # Model identifier in LM Studio
timeout_seconds = 180                        # Per-request LLM timeout (seconds)
temperature     = 0.1                        # Low temp → more deterministic output
max_output_tokens = 2048                     # Keep ≤ 2048 to prevent hangs

[pipeline]
batch_size            = 3    # Pages per LLM call
sparse_page_threshold = 30   # Min characters to consider a page non-empty
extra_slot_count      = 10   # Extra_ columns available (extra_1…extra_10)
pdftotext_path        = bin/pdftotext.exe
```

---

## Usage

### Process a single PDF

```powershell
python pipeline.py path\to\invoice.pdf
```

#### Outputs

For each processed file `sample.pdf`, the pipeline produces up to 5 files in the same directory:

| File | Description |
|------|-------------|
| `<name>.xlsx` | Final structured Excel spreadsheet conforming to the fixed schema |
| `<name>_raw.xlsx` | Unconverted Excel preserving raw model values and extra discovered keys |
| `<name>_hardening_report.json` | Diagnostic report with type warnings, invented keys, and execution metrics |
| `<name>_response.json` | Full LM Studio request prompt and raw model JSON completion |
| `<name>_extracted.txt` | Clean text extracted from the document via pdftotext |

### Process a plain-text file

```powershell
python pipeline.py path\to\invoice.txt
```

---

## Testing

### Unit tests (fast, no LLM required)

```powershell
python -m unittest tests\test_pipeline.py -v
```

All 15 unit tests cover: `convert_number`, `convert_date`, `parse_response`,
`extraction_quality_issues`, `has_usable_extraction`, and `flatten_fixed`.

### Full integration test runner (requires LM Studio running)

```powershell
# Run all 15 sample invoices
python test_runner.py

# Run a specific subset
python test_runner.py --samples 1,3,5,8

# Set a per-sample timeout (default 300 s)
python test_runner.py --timeout 300
```

Results are written to:

- `test_files/milestone7_test_summary.csv`
- `TEST_REPORT.md`

### PowerShell batch runner

```powershell
.\test_all_15_samples.ps1
```

---

## Project Structure

```
document_excel_pipeline/
├── pipeline.py                 # Main pipeline (single-file, ~1916 lines)
├── config.ini                  # Active configuration
├── config.example.ini          # Config template with comments
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── HANDOVER.md                 # Internship handover guide
├── TEST_REPORT.md              # Full test results (15 samples)
├── test_runner.py              # Reproducible integration test runner
├── test_all_15_samples.ps1     # PowerShell batch runner for all samples
├── bin/
│   └── pdftotext.exe           # Poppler text extractor (Windows)
├── test_files/
│   ├── sample_invoice.pdf      # Sample 1
│   ├── sample_invoice2.pdf     # Samples 2–15
│   └── …
└── tests/
    └── test_pipeline.py        # 15 unit tests
```

---

## Output Excel Schema

### Header row (one row per document)

| Column | Description |
|--------|-------------|
| `document_type` | Invoice / Quote / Delivery Note / etc. |
| `document_number` | Invoice / PO number |
| `document_date` | Issue date (ISO 8601 where possible) |
| `due_date` | Payment due date |
| `seller_name` | Vendor / supplier name |
| `seller_address` | Vendor address |
| `seller_contact` | Vendor phone / email |
| `seller_tax_id` | GST / VAT / ABN / EIN |
| `buyer_name` | Customer name |
| `buyer_address` | Customer address |
| `buyer_contact` | Customer phone / email |
| `buyer_tax_id` | Customer tax ID |
| `currency` | Currency code (USD, INR, AUD, …) |
| `total_amount` | Grand total / invoice total |
| `tax_amount` | GST / VAT / tax total |
| `reference_number` | PO number / reference |
| `notes` | Terms, payment instructions, free-text |

### Line items (one row per line item)

| Column | Description |
|--------|-------------|
| `item_description` | Product / service name |
| `quantity` | Units |
| `unit` | Unit of measure |
| `rate` | Unit price |
| `line_amount` | Extended price |
| `item_tax` | Line-level tax |

### Extra columns

`extra_1` … `extra_10` — overflow fields for any data the LLM finds that
does not fit the fixed schema.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `ConnectionError` / timeout | LM Studio not running | Start LM Studio and load the model |
| Excel not generated | Validation failed after 3 LLM retries | Check `_pipeline_report.json` for the error detail |
| All fields blank | PDF is scanned (image-only) | OCR is out of scope — use a text-based PDF |
| Date appears as text | Unusual date format | Check `convert_date()` in `pipeline.py`; add the format string |
| `RuntimeError: Batch N: extraction failed` | Model hallucinated badly 3× | Try a larger model in LM Studio (`qwen2.5-7b-instruct`) |

---

## Known Limitations

- **Scanned / image PDFs**: `pdftotext` extracts no text from image-only pages. OCR integration is out of scope for this version.
- **Word-form amounts**: Values like *"Seven Thousand Five Hundred"* are kept as-is (not converted to numbers).
- **Calculated amounts**: If a document shows a unit price but the extended amount is calculated (e.g., after a discount), the LLM may compute the pre-discount figure. The pipeline logs a warning but still writes the extracted value.
- **Model non-determinism**: At `temperature=0.1`, output is mostly stable but can vary between runs.
- **Windows only**: Poppler binaries in `bin/` are Windows x64 executables.

---

## License

Internal project — not for public distribution.
