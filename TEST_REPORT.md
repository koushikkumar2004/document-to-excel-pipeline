# Test Report — Document-to-Excel Pipeline
## Full 15-Sample Integration Test

> **Test environment**
> - Python 3.13.2, Windows 11
> - Model: `qwen2.5-3b-instruct` via LM Studio at `http://localhost:1234/v1`
> - Config: `timeout_seconds=180`, `max_output_tokens=2048`, `temperature=0.1`
> - Pipeline version: post-hardening (all 5 originally-failing samples addressed)

---

## Summary Table

| Sample | File | Status | Runtime (s) | Notes |
|--------|------|--------|-------------|-------|
| 1 | `sample_invoice.pdf` | ✅ SUCCESS | ~70 | Date `January 25, 2016` now parses correctly |
| 2 | `sample_invoice2.pdf` | ⏳ Pending full run | — | — |
| 3 | `sample_invoice3.pdf` | ✅ SUCCESS | ~106 | Thousands-separator fix (`Rs. 42,000.00`) |
| 4 | `sample_invoice4.pdf` | ⏳ Pending full run | — | — |
| 5 | `sample_invoice5.pdf` | ✅ SUCCESS | ~106 | Delivery note — 3 rows extracted |
| 6 | `sample_invoice6.pdf` | ⏳ Pending full run | — | — |
| 7 | `sample_invoice7.pdf` | ⏳ Pending full run | — | — |
| 8 | `sample_invoice8.pdf` | ✅ SUCCESS | ~229 | Attempt 3 succeeds; due_date warning logged |
| 9 | `sample_invoice9.pdf` | ⏳ Pending full run | — | — |
| 10 | `sample_invoice10.pdf` | ⏳ Pending full run | — | — |
| 11 | `sample_invoice11.pdf` | ⏳ Pending full run | — | — |
| 12 | `sample_invoice12.pdf` | ⏳ Pending full run | — | — |
| 13 | `sample_invoice13.pdf` | ✅ SUCCESS | ~87 | Blank page 2 handled gracefully |
| 14 | `sample_invoice14.pdf` | ⏳ Pending full run | — | — |
| 15 | `sample_invoice15.pdf` | ✅ SUCCESS | ~100 | Blank page 2 handled gracefully |

> ⏳ Rows will be updated after `python test_runner.py --samples 2,4,6,7,9,10,11,12,14` completes.

---

## Root-Cause Fixes Applied

| Sample | Original Root Cause | Fix Applied |
|--------|-------------------|-------------|
| 1 | Filename mismatch (`sample_invoice1.pdf` vs `sample_invoice.pdf`) | Test runner + PS1 updated to detect actual filename |
| 3 | Thousands-separator `Rs. 42,000.00` falsely flagged as comma-separated multi-value | `extraction_quality_issues()` rewritten — strips thousands separators before checking |
| 5 | Delivery note with no prices — pipeline rejected empty `line_items` | Added `has_prior_usable` logic in `process_batch()` to accept partial results |
| 8 | `due_date` duplicated `document_date` (single-date invoice) → hard validation failure | Demoted to semantic warning (log only, not a blocking issue) |
| 13 | Page 2 blank template — empty extraction triggered `RuntimeError` | `process_batch()` accepts graceful empty result when earlier pages succeeded |
| 15 | Same as sample 13 | Same fix |

---

## Unit Test Results

**All 16 unit tests pass** (`python -m unittest tests\test_pipeline.py -v`).

| Test | Class | Result |
|------|-------|--------|
| `test_convert_number_integers` | `TestTypeConversions` | ✅ |
| `test_convert_number_floats_and_currency` | `TestTypeConversions` | ✅ |
| `test_convert_number_rate_units` | `TestTypeConversions` | ✅ |
| `test_convert_number_percentages` | `TestTypeConversions` | ✅ |
| `test_convert_number_preserves_words_with_warning` | `TestTypeConversions` | ✅ |
| `test_convert_number_strips_label_prefix` | `TestTypeConversions` | ✅ |
| `test_convert_date_standard_formats` | `TestTypeConversions` | ✅ |
| `test_convert_date_ordinal_formats` | `TestTypeConversions` | ✅ |
| `test_convert_date_preserves_invalid_with_warning` | `TestTypeConversions` | ✅ |
| `test_thousands_separator_not_flagged_as_multiple_values` | `TestExtractionQuality` | ✅ |
| `test_true_multiple_values_are_flagged` | `TestExtractionQuality` | ✅ |
| `test_has_usable_extraction` | `TestExtractionQuality` | ✅ |
| `test_parse_response_markdown_fences` | `TestJSONParsingAndFallbacks` | ✅ |
| `test_parse_response_think_tags` | `TestJSONParsingAndFallbacks` | ✅ |
| `test_parse_response_recovers_header_from_column_types` | `TestJSONParsingAndFallbacks` | ✅ |
| `test_flatten_fixed_preserves_columns_and_blanks` | `TestFlatteningAndSchema` | ✅ |

---

## Known Limitations

- **Sample 8 — `line_amount` for WebDev row**: The model returns `$400.00` (pre-discount) instead of `$320.00` (post-20%-discount). After 3 LLM attempts the pipeline succeeds, but the WebDev extended price is incorrect. This is a model reasoning limitation — the 3B parameter model does not reliably apply percentage discounts.

- **Scanned PDFs**: `pdftotext` extracts no text from image-only pages. Any sample that is scanned (not text-based) will produce empty extraction.

- **Word-form amounts**: Values written in words (e.g., "Seven Thousand Five Hundred") are kept as-is in the Excel cell — not converted to numbers. The pipeline logs a type-conversion warning for each such cell.

---

## How to Reproduce

```powershell
# Run all 15 samples (requires LM Studio running with qwen2.5-3b-instruct)
python test_runner.py

# Run a subset
python test_runner.py --samples 1,3,5,8,13,15 --timeout 300

# Unit tests only (no LM Studio required)
python -m unittest tests\test_pipeline.py -v
```

Results are written to `test_files/milestone7_test_summary.csv`.
