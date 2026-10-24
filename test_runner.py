#!/usr/bin/env python3
"""
Reproducible Test Runner for Document-to-Excel Automation Pipeline.

Features:
- Discovers actual sample filenames (handles sample_invoice.pdf vs sample_invoice1.pdf).
- Cleans target test's previous generated outputs before running to avoid stale Excel false positives.
- Executes pipeline with bounded process timeout.
- Collects execution time, exit codes, hardening diagnostics (invented keys, warnings, dropped columns).
- Evaluates extraction accuracy against verifiable ground-truth values from source documents.
- Produces CSV summary (test_files/test_summary.csv) and Markdown report (TEST_REPORT.md).
- Distinguishes execution success from extraction accuracy.
"""

import argparse
import csv
import json
import os
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path
import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent
TEST_DIR = PROJECT_ROOT / "test_files"
PIPELINE_SCRIPT = PROJECT_ROOT / "pipeline.py"

# Ground-truth verifiable values extracted directly from source documents
GROUND_TRUTH = {
    1: {
        "name": "sample_invoice.pdf",
        "doc_number": "INV-3337",
        "doc_date": date(2016, 1, 25),
        "total_amount": 93.5,
        "seller": "DEMO - Sliced Invoices",
        "buyer": "Test Business",
        "min_rows": 1,
    },
    2: {
        "name": "sample_invoice2.pdf",
        "doc_number": "112456",
        "total_amount": 64.93,
        "seller": "Vortex Services",
        "min_rows": 1,
    },
    3: {
        "name": "sample_invoice3.pdf",
        "doc_number": "BR-2026-00001",
        "doc_date": date(2026, 9, 8),
        "total_amount": 54870.0,
        "seller": "Sharma Computers",
        "buyer": "Mysuru Design Studio",
        "min_rows": 2,
    },
    4: {
        "name": "sample_invoice4.pdf",
        "doc_number": "3872",
        "total_amount": 4220.0,
        "min_rows": 4,
    },
    5: {
        "name": "sample_invoice5.pdf",
        "doc_number": "DN-8821",
        "doc_date": date(2026, 3, 25),
        "seller": "Northline Press",
        "buyer": "Harborpoint warehouse",
        "min_rows": 3,
    },
    6: {
        "name": "sample_invoice6.pdf",
        "doc_number": "6489",
        "total_amount": 1888.0,
        "min_rows": 1,
    },
    7: {
        "name": "sample_invoice7.pdf",
        "doc_number": "INV1048",
        "total_amount": 1625.0,
        "min_rows": 9,
    },
    8: {
        "name": "sample_invoice8.pdf",
        "doc_number": "0001",
        "subtotal_amount": 495.0,
        "doc_date": date(2003, 5, 6),
        "seller": "Your Company",
        "buyer": "Client Name",
        "min_rows": 3,
    },
    9: {
        "name": "sample_invoice9.pdf",
        "doc_number": "112456",
        "total_amount": 64.93,
        "min_rows": 1,
    },
    10: {
        "name": "sample_invoice10.pdf",
        "doc_number": "INV-12345",
        "total_amount": 1150.0,
        "min_rows": 1,
    },
    11: {
        "name": "sample_invoice11.pdf",
        "doc_number": "1001",
        "total_amount": 107.0,
        "min_rows": 1,
    },
    12: {
        "name": "sample_invoice12.pdf",
        "doc_number": "99824",
        "total_amount": 22800.0,
        "min_rows": 20,
    },
    13: {
        "name": "sample_invoice13.pdf",
        "doc_number": "8752",
        "doc_date": date(2024, 5, 22),
        "seller": "CHUNNU LAL BAKERY",
        "min_rows": 1,
    },
    14: {
        "name": "sample_invoice14.pdf",
        "doc_number": "INV-10012",
        "total_amount": 1250.0,
        "min_rows": 1,
    },
    15: {
        "name": "sample_invoice15.pdf",
        "doc_number": "244",
        "doc_date": date(2024, 5, 25),
        "seller": "CHUNNU LAL BAKERY",
        "total_amount": 7000.0,
        "min_rows": 1,
    },
}


def discover_sample_file(sample_id):
    """Resolve sample id to actual existing file path."""
    if sample_id == 1:
        candidates = [TEST_DIR / "sample_invoice.pdf", TEST_DIR / "sample_invoice1.pdf"]
    else:
        candidates = [TEST_DIR / f"sample_invoice{sample_id}.pdf"]

    for cand in candidates:
        if cand.exists():
            return cand
    return None


def clean_target_outputs(input_file):
    """Safely remove previous outputs only for this specific sample to prevent stale reads."""
    stem = input_file.stem
    parent = input_file.parent
    targets = [
        parent / f"{stem}.xlsx",
        parent / f"{stem}_raw.xlsx",
        parent / f"{stem}_hardening_report.json",
    ]
    for target in targets:
        if target.exists():
            try:
                target.unlink()
            except OSError as exc:
                print(f"Warning: Could not remove old file {target}: {exc}")


def evaluate_accuracy(excel_path, expected):
    """
    Compare generated Excel file against verifiable expected values.
    Returns (status, passed_checks, failed_checks, total_checks).
    """
    if not excel_path.exists():
        return "NO_OUTPUT", [], ["Excel file missing"], 0

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
    except Exception as exc:
        return "READ_ERROR", [], [f"Cannot read Excel: {exc}"], 0

    # Header row mapping
    headers = [cell.value for cell in ws[1]]
    col_map = {name: idx + 1 for idx, name in enumerate(headers) if name}

    passed = []
    failed = []

    def get_val(row_idx, col_name):
        col_idx = col_map.get(col_name)
        if not col_idx:
            return None
        return ws.cell(row=row_idx, column=col_idx).value

    # Check minimum rows
    row_count = ws.max_row - 1  # exclude header
    min_rows = expected.get("min_rows", 1)
    if row_count >= min_rows:
        passed.append(f"Row count ({row_count} >= {min_rows})")
    else:
        failed.append(f"Row count ({row_count} < expected {min_rows})")

    # Document Number
    if "doc_number" in expected:
        expected_num = str(expected["doc_number"]).strip()
        actual_num = str(get_val(2, "document_number") or "").strip()
        if expected_num.lower() in actual_num.lower():
            passed.append(f"Document Number '{expected_num}'")
        else:
            failed.append(f"Document Number (expected '{expected_num}', got '{actual_num}')")

    # Document Date
    if "doc_date" in expected:
        exp_date = expected["doc_date"]
        actual_date = get_val(2, "document_date")
        if isinstance(actual_date, (datetime, date)):
            act_date = actual_date.date() if isinstance(actual_date, datetime) else actual_date
            if act_date == exp_date:
                passed.append(f"Document Date {exp_date}")
            else:
                failed.append(f"Document Date (expected {exp_date}, got {act_date})")
        else:
            failed.append(f"Document Date (expected {exp_date}, got {actual_date})")

    # Total Amount
    if "total_amount" in expected:
        exp_amt = expected["total_amount"]
        actual_amt = get_val(2, "total_amount")
        try:
            if actual_amt is not None and abs(float(actual_amt) - exp_amt) < 0.01:
                passed.append(f"Total Amount {exp_amt}")
            else:
                failed.append(f"Total Amount (expected {exp_amt}, got {actual_amt})")
        except (ValueError, TypeError):
            failed.append(f"Total Amount (expected {exp_amt}, got non-numeric '{actual_amt}')")

    # Seller Name
    if "seller" in expected:
        exp_seller = expected["seller"].lower()
        actual_seller = str(get_val(2, "seller_name") or "").lower()
        if exp_seller in actual_seller or actual_seller in exp_seller:
            passed.append(f"Seller '{expected['seller']}'")
        else:
            failed.append(f"Seller (expected '{expected['seller']}', got '{get_val(2, 'seller_name')}')")

    # Buyer Name
    if "buyer" in expected:
        exp_buyer = expected["buyer"].lower()
        actual_buyer = str(get_val(2, "buyer_name") or "").lower()
        if exp_buyer in actual_buyer or actual_buyer in exp_buyer:
            passed.append(f"Buyer '{expected['buyer']}'")
        else:
            failed.append(f"Buyer (expected '{expected['buyer']}', got '{get_val(2, 'buyer_name')}')")

    total_checks = len(passed) + len(failed)
    if not failed:
        status = "PASSED"
    elif passed:
        status = "PARTIAL"
    else:
        status = "FAILED"

    return status, passed, failed, total_checks


def run_single_test(sample_id, timeout_sec=300):
    """Execute pipeline on a single sample and collect detailed test results."""
    pdf_path = discover_sample_file(sample_id)
    if not pdf_path:
        return {
            "sample": sample_id,
            "filename": f"sample_invoice{sample_id}.pdf",
            "execution_status": "MISSING_FILE",
            "duration": 0.0,
            "excel_path": "",
            "rows": 0,
            "invented_keys": 0,
            "type_warnings": 0,
            "dropped_columns": 0,
            "accuracy_status": "NOT_EVALUATED",
            "passed_checks": [],
            "failed_checks": ["Source document not found"],
            "error": "File does not exist",
        }

    stem = pdf_path.stem
    excel_path = pdf_path.parent / f"{stem}.xlsx"
    report_path = pdf_path.parent / f"{stem}_hardening_report.json"

    # Clean old target output to ensure freshness
    clean_target_outputs(pdf_path)

    start_time = time.time()
    try:
        proc = subprocess.run(
            [sys.executable, str(PIPELINE_SCRIPT), str(pdf_path)],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_sec,
        )
        duration = round(time.time() - start_time, 2)
        exit_code = proc.returncode
        error_msg = proc.stderr.strip() if proc.stderr else ""
    except subprocess.TimeoutExpired:
        duration = round(time.time() - start_time, 2)
        exit_code = -1
        error_msg = f"Timed out after {timeout_sec} seconds"
    except Exception as exc:
        duration = round(time.time() - start_time, 2)
        exit_code = -1
        error_msg = str(exc)

    excel_exists = excel_path.exists()
    execution_status = "SUCCESS" if exit_code == 0 and excel_exists else "FAILED"

    # Read hardening report if available
    invented_keys = 0
    type_warnings = 0
    dropped_columns = 0
    row_count = 0
    if report_path.exists():
        try:
            data = json.loads(report_path.read_text(encoding="utf-8"))
            invented_keys = len(data.get("invented_keys", []))
            type_warnings = len(data.get("type_warnings", []))
            dropped_columns = len(data.get("dropped_columns", []))
            row_count = data.get("excel_row_count", 0)
        except Exception:
            pass

    # Accuracy checks
    acc_status, passed_checks, failed_checks, total_checks = (
        "NOT_EVALUATED",
        [],
        [],
        0,
    )
    if execution_status == "SUCCESS" and sample_id in GROUND_TRUTH:
        acc_status, passed_checks, failed_checks, total_checks = evaluate_accuracy(
            excel_path, GROUND_TRUTH[sample_id]
        )

    return {
        "sample": sample_id,
        "filename": pdf_path.name,
        "execution_status": execution_status,
        "duration": duration,
        "excel_path": str(excel_path) if excel_exists else "",
        "rows": row_count,
        "invented_keys": invented_keys,
        "type_warnings": type_warnings,
        "dropped_columns": dropped_columns,
        "accuracy_status": acc_status,
        "passed_checks": passed_checks,
        "failed_checks": failed_checks,
        "error": error_msg,
    }


def write_csv_summary(results, csv_path):
    """Write standard execution and accuracy summary CSV."""
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "Sample",
                "Filename",
                "ExecutionStatus",
                "DurationSeconds",
                "ExcelPath",
                "Rows",
                "InventedKeys",
                "TypeWarnings",
                "DroppedColumns",
                "AccuracyStatus",
                "AccuracyChecksPassed",
                "AccuracyChecksFailed",
                "FailureReason",
            ]
        )
        for r in results:
            writer.writerow(
                [
                    r["sample"],
                    r["filename"],
                    r["execution_status"],
                    r["duration"],
                    r["excel_path"],
                    r["rows"],
                    r["invented_keys"],
                    r["type_warnings"],
                    r["dropped_columns"],
                    r["accuracy_status"],
                    "; ".join(r["passed_checks"]),
                    "; ".join(r["failed_checks"]),
                    r["error"],
                ]
            )


def write_markdown_report(results, report_path):
    """Write human-readable comprehensive test report."""
    total = len(results)
    exec_success = sum(1 for r in results if r["execution_status"] == "SUCCESS")
    exec_failed = sum(1 for r in results if r["execution_status"] != "SUCCESS")
    acc_passed = sum(1 for r in results if r["accuracy_status"] == "PASSED")
    acc_partial = sum(1 for r in results if r["accuracy_status"] == "PARTIAL")
    total_time = sum(r["duration"] for r in results)

    content = [
        "# Document-to-Excel Pipeline Test Report",
        "",
        f"**Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Executive Summary",
        "",
        f"- **Total Samples Evaluated:** {total}",
        f"- **Execution Success Rate:** {exec_success}/{total} ({exec_success/total*100:.1f}%)",
        f"- **Execution Failures:** {exec_failed}",
        f"- **Extraction Accuracy Pass Rate:** {acc_passed}/{total} fully passed, {acc_partial}/{total} partially verified",
        f"- **Total Test Duration:** {total_time:.2f} seconds ({total_time / 60:.2f} minutes)",
        "",
        "## Per-Sample Results",
        "",
        "| Sample | File | Execution | Runtime (s) | Rows | Warnings | Accuracy | Notes / Verifications |",
        "|:---:|:---|:---:|---:|:---:|:---:|:---:|:---|",
    ]

    for r in results:
        passed_str = f"{len(r['passed_checks'])} passed" if r['passed_checks'] else "-"
        fails_str = f", {len(r['failed_checks'])} failed" if r['failed_checks'] else ""
        notes = f"{passed_str}{fails_str}"
        if r["error"]:
            notes += f" | Error: {r['error'][:60]}"
        content.append(
            f"| {r['sample']} | `{r['filename']}` | **{r['execution_status']}** | {r['duration']} | "
            f"{r['rows']} | {r['type_warnings']} | {r['accuracy_status']} | {notes} |"
        )

    content.extend(
        [
            "",
            "## Accuracy Check Details",
            "",
        ]
    )

    for r in results:
        content.append(f"### Sample {r['sample']}: `{r['filename']}`")
        content.append(f"- **Execution Status:** {r['execution_status']}")
        content.append(f"- **Runtime:** {r['duration']}s")
        if r["passed_checks"]:
            content.append("- **Verified Correct Checks:**")
            for chk in r["passed_checks"]:
                content.append(f"  - [x] {chk}")
        if r["failed_checks"]:
            content.append("- **Failed / Unmatched Checks:**")
            for chk in r["failed_checks"]:
                content.append(f"  - [ ] {chk}")
        content.append("")

    report_path.write_text("\n".join(content), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Run pipeline tests on sample invoices.")
    parser.add_argument(
        "--samples",
        default="all",
        help="Comma-separated sample numbers (e.g. 1,2,3 or 1-5) or 'all' (default)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=300,
        help="Max timeout per sample in seconds (default: 300)",
    )
    parser.add_argument(
        "--csv",
        default=str(TEST_DIR / "milestone7_test_summary.csv"),
        help="CSV summary output path",
    )
    parser.add_argument(
        "--report",
        default=str(PROJECT_ROOT / "TEST_REPORT.md"),
        help="Markdown report output path",
    )
    args = parser.parse_args()

    if args.samples == "all":
        sample_ids = list(range(1, 16))
    else:
        sample_ids = []
        for part in args.samples.split(","):
            part = part.strip()
            if "-" in part:
                s, e = part.split("-", 1)
                sample_ids.extend(range(int(s), int(e) + 1))
            else:
                sample_ids.append(int(part))

    print("=" * 70)
    print("DOCUMENT-TO-EXCEL AUTOMATION TEST RUNNER")
    print("=" * 70)
    print(f"Target Samples: {sample_ids}")
    print(f"Per-sample Timeout: {args.timeout}s")
    print(f"CSV Summary: {args.csv}")
    print(f"Markdown Report: {args.report}")
    print()

    results = []
    suite_start = time.time()

    for idx, sid in enumerate(sample_ids, start=1):
        print(f"[{idx}/{len(sample_ids)}] Running Sample {sid}...")
        r = run_single_test(sid, timeout_sec=args.timeout)
        results.append(r)
        print(
            f"       Execution: {r['execution_status']} in {r['duration']}s | "
            f"Accuracy: {r['accuracy_status']} | Rows: {r['rows']} | Warnings: {r['type_warnings']}"
        )
        print()

    total_elapsed = round(time.time() - suite_start, 2)
    print("=" * 70)
    print(f"TEST SUITE COMPLETE in {total_elapsed}s ({total_elapsed/60:.2f} mins)")
    print("=" * 70)

    # Save outputs
    write_csv_summary(results, Path(args.csv))
    write_markdown_report(results, Path(args.report))
    print(f"Summary CSV saved to: {args.csv}")
    print(f"Markdown report saved to: {args.report}")


if __name__ == "__main__":
    main()
