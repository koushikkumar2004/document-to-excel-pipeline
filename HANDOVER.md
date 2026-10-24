# Internship Handover Guide — Document-to-Excel Pipeline

**Project**: `document_excel_pipeline`  
**Handover date**: October 2026  
**Status**: Functional — 14/15 sample invoices produce correct Excel output.

---

## 1. What This Project Does

This tool automates the conversion of PDF invoices into structured Excel
spreadsheets. You point it at a PDF; it uses a local AI model (via LM Studio)
to read the invoice and fill in a fixed spreadsheet schema.

No cloud API keys are needed. Everything runs on your local machine.

---

## 2. Quick-Start Checklist

Complete these steps in order before running anything:

- [ ] Install **Python 3.10+** (tested with 3.13.2)
- [ ] Run `pip install -r requirements.txt`
- [ ] Install **LM Studio** from https://lmstudio.ai/
- [ ] In LM Studio, download model **`qwen2.5-3b-instruct`**
- [ ] In LM Studio, open the **Local Server** tab and start the server
- [ ] Load the Qwen model (the server status should say "Model loaded")
- [ ] Run a quick test: `python pipeline.py test_files/sample_invoice.pdf`
- [ ] Check the output: `test_files/sample_invoice_output.xlsx` should appear

---

## 3. How to Run

### Single invoice
```powershell
python pipeline.py path\to\invoice.pdf
```

The Excel file and a JSON diagnostic report are written next to the PDF.

### Batch test all 15 samples
```powershell
python test_runner.py
```

Results appear in `TEST_REPORT.md` and `test_files/milestone7_test_summary.csv`.

---

## 4. Folder Structure (what matters)

```
document_excel_pipeline/
├── pipeline.py          ← THE pipeline (all logic lives here, ~1916 lines)
├── config.ini           ← Settings: LM Studio URL, model name, timeouts
├── requirements.txt     ← pip dependencies
├── README.md            ← Full documentation
├── test_runner.py       ← Integration test runner
├── tests/
│   └── test_pipeline.py ← 15 unit tests (run without LM Studio)
├── bin/
│   └── pdftotext.exe    ← PDF text extractor — do NOT delete
└── test_files/          ← 15 sample invoices + their Excel outputs
```

---

## 5. Configuration Reference

Open `config.ini`. The most important settings:

| Setting | Default | When to change |
|---------|---------|----------------|
| `base_url` | `http://localhost:1234/v1` | If LM Studio uses a different port |
| `model` | `qwen2.5-3b-instruct` | If you switch to a larger Qwen model |
| `timeout_seconds` | `180` | Increase if your machine is slow |
| `max_output_tokens` | `2048` | Keep ≤ 2048 — higher values cause hangs |
| `temperature` | `0.1` | Keep low for deterministic extraction |

---

## 6. Running the Tests

### Unit tests (fast, ~3 seconds, no LM Studio needed)
```powershell
python -m unittest tests\test_pipeline.py -v
```
All 15 should pass. These test number conversion, date parsing, JSON parsing,
and validation logic.

### Integration tests (slow, ~2-5 minutes each, needs LM Studio)
```powershell
python test_runner.py --samples 1,3,5,8 --timeout 300
```

---

## 7. Understanding the Code

`pipeline.py` is one self-contained file. Key sections:

| Lines | What it does |
|-------|-------------|
| 1–130 | Config loading, schema constants (`HEADER_COLUMNS`, `LINE_ITEM_COLUMNS`) |
| 253–403 | `extract_pdf()` — calls `pdftotext.exe`, splits pages |
| 411–615 | `build_prompt()` — builds the JSON extraction prompt sent to the LLM |
| 620–685 | `request_model()` — HTTP POST to LM Studio |
| 687–800 | `parse_response()` — cleans and validates the LLM's JSON reply |
| 942–1040 | `extraction_quality_issues()` — checks the extraction for structural problems |
| 1069–1131 | `process_batch()` — retries the LLM up to 3× per batch |
| 1345–1465 | `convert_number()`, `convert_date()` — type normalisation |
| 1600–1916 | `main()`, `save_workbook()` — Excel generation and report writing |

### Retry logic
For each batch of pages, the pipeline tries up to 3 prompts:
1. **Normal prompt** — full schema template
2. **Recovery prompt** — stricter instructions, includes the previous bad output
3. **Table-recovery prompt** — simplified, asks for plain tabular output

If all 3 fail, the pipeline raises `RuntimeError` and **no Excel is written**.

---

## 8. Known Issue — Sample 8

`sample_invoice8.pdf` is a multi-section invoice with a 20% discount row.
The LLM calculates the pre-discount line amount (\$400) instead of reading the
printed post-discount total (\$320). This is a model reasoning limitation, not
a code bug.

**Current behaviour**: The pipeline succeeds, but `line_amount` for the
WebDev row will show `$400` instead of `$320`. A semantic warning is logged.

**If you want to fix it properly**: Either switch to a larger model
(`qwen2.5-7b-instruct`) or post-process the line amount by checking the
`line_amount` against numbers literally present in the PDF text.

---

## 9. Adding a New Invoice Type

The pipeline is general-purpose. For a new invoice:
1. Place the PDF in `test_files/` (or any path).
2. Run `python pipeline.py path\to\invoice.pdf`.
3. Inspect `_pipeline_report.json` if the output looks wrong.
4. If a field is being missed, check whether the field name is in
   `HEADER_COLUMNS` (line ~90 of `pipeline.py`). If not, use an `extra_N` column.

---

## 10. Common Errors and Fixes

| Error | Fix |
|-------|-----|
| `ConnectionError: http://localhost:1234` | LM Studio is not running — start it and load the model |
| `RuntimeError: Batch N: extraction failed after 3 attempts` | Try a larger Qwen model in LM Studio, or simplify the invoice |
| Excel is blank / all empty cells | The PDF is image-only (scanned) — pdftotext extracts nothing from images |
| Date shows as `"January 25, 2016"` (text, not date) | This is handled — but check `convert_date()` if a new format appears |
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |

---

## 11. Git History

The repo tracks all changes from the baseline:

```powershell
git log --oneline
```

The baseline commit is tagged **"Baseline: restored working version"** so you
can always diff against it:

```powershell
git diff HEAD~1 pipeline.py
```

---

## 12. Next Steps / Suggestions for Future Developers

- **OCR support**: Integrate Tesseract or AWS Textract for scanned PDFs.
- **Larger model**: `qwen2.5-7b-instruct` gives more accurate extraction but is slower.
- **Web UI**: Wrap `pipeline.py` in a simple Flask endpoint for drag-and-drop uploads.
- **Batch input folder**: Add a CLI flag `--folder` to process an entire directory of PDFs.
- **Confidence scores**: Ask the LLM to return a confidence field per extraction and flag low-confidence cells in the Excel.

---

*For technical questions about the code, start with the README and the inline
comments in `pipeline.py`. The unit tests in `tests/test_pipeline.py` are also
a good reference for how each function behaves.*
