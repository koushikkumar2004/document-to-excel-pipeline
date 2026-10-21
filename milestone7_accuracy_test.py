import json
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime


# ============================================================
# CONFIGURATION
# ============================================================

PROJECT_DIR = Path(".")

TEST_DIR = PROJECT_DIR / "test_files"

PIPELINE_SCRIPT = PROJECT_DIR / "pipeline.py"

RESULT_FILE = TEST_DIR / "milestone7_accuracy_results.json"

START_SAMPLE = 1
END_SAMPLE = 15


# ============================================================
# HELPERS
# ============================================================

def find_extracted_file(sample_number):
    """
    Find the extracted text file for a sample.
    """

    extracted_file = (
        TEST_DIR
        / f"sample_invoice{sample_number}_extracted.txt"
    )

    if extracted_file.exists():
        return extracted_file

    # sample_invoice.pdf is special because it has
    # no number in its filename.
    if sample_number == 1:

        extracted_file = (
            TEST_DIR
            / "sample_invoice_extracted.txt"
        )

        if extracted_file.exists():
            return extracted_file

    return None


def find_pdf_file(sample_number):
    """
    Find the original PDF.
    """

    pdf_file = (
        TEST_DIR
        / f"sample_invoice{sample_number}.pdf"
    )

    if pdf_file.exists():
        return pdf_file

    if sample_number == 1:

        pdf_file = (
            TEST_DIR
            / "sample_invoice.pdf"
        )

        if pdf_file.exists():
            return pdf_file

    return None


def count_pages_from_extracted(file_path):

    text = file_path.read_text(
        encoding="utf-8-sig",
        errors="replace"
    )

    return text.count(
        "--- Page "
    )


# ============================================================
# RUN ONE SAMPLE
# ============================================================

def run_sample(
    sample_number,
    input_file
):

    print()
    print("=" * 70)
    print(
        f"SAMPLE {sample_number}"
    )
    print("=" * 70)

    print(
        f"Input: {input_file}"
    )

    start_time = time.time()

    try:

        result = subprocess.run(
            [
                sys.executable,
                str(PIPELINE_SCRIPT),
                str(input_file)
            ],
            cwd=PROJECT_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )

        duration = (
            time.time()
            - start_time
        )

        success = (
            result.returncode == 0
        )

        print(
            f"Duration: "
            f"{duration:.2f} seconds"
        )

        print(
            f"Exit code: "
            f"{result.returncode}"
        )

        if success:

            print(
                "STATUS: SUCCESS"
            )

        else:

            print(
                "STATUS: FAILED"
            )

        return {
            "sample_number": sample_number,
            "input_file": str(input_file),
            "duration_seconds": round(
                duration,
                2
            ),
            "exit_code": result.returncode,
            "status": (
                "SUCCESS"
                if success
                else "FAILED"
            ),
            "stdout": result.stdout,
            "stderr": result.stderr
        }

    except Exception as error:

        duration = (
            time.time()
            - start_time
        )

        print(
            f"STATUS: ERROR"
        )

        print(
            f"Error: {error}"
        )

        return {
            "sample_number": sample_number,
            "input_file": str(input_file),
            "duration_seconds": round(
                duration,
                2
            ),
            "exit_code": -1,
            "status": "ERROR",
            "stdout": "",
            "stderr": str(error)
        }


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print(
        "MILESTONE 7 - ACCURACY TEST"
    )
    print("=" * 70)

    start_time = time.time()

    start_timestamp = (
        datetime.now()
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    print(
        f"Start time: {start_timestamp}"
    )

    print(
        f"Testing samples "
        f"{START_SAMPLE} to {END_SAMPLE}"
    )

    print()

    # --------------------------------------------------------
    # Check pipeline
    # --------------------------------------------------------

    if not PIPELINE_SCRIPT.exists():

        print(
            f"ERROR: {PIPELINE_SCRIPT} "
            f"was not found."
        )

        print(
            "Create/verify pipeline.py first."
        )

        sys.exit(1)

    # --------------------------------------------------------
    # Find samples
    # --------------------------------------------------------

    samples = []

    for sample_number in range(
        START_SAMPLE,
        END_SAMPLE + 1
    ):

        pdf_file = find_pdf_file(
            sample_number
        )

        extracted_file = find_extracted_file(
            sample_number
        )

        if pdf_file is None:

            print(
                f"WARNING: sample {sample_number} "
                f"PDF not found."
            )

        if extracted_file is None:

            print(
                f"WARNING: sample {sample_number} "
                f"extracted TXT not found."
            )

        if (
            pdf_file is not None
            and extracted_file is not None
        ):

            samples.append(
                {
                    "sample_number": sample_number,
                    "pdf": pdf_file,
                    "extracted": extracted_file
                }
            )

    print(
        f"Samples available: "
        f"{len(samples)}"
    )

    print()

    # --------------------------------------------------------
    # Run samples
    # --------------------------------------------------------

    results = []

    for index, sample in enumerate(
        samples,
        start=1
    ):

        sample_number = (
            sample["sample_number"]
        )

        input_file = sample["pdf"]

        print(
            f"Progress: "
            f"{index}/{len(samples)}"
        )

        result = run_sample(
            sample_number,
            input_file
        )

        result[
            "page_count"
        ] = count_pages_from_extracted(
            sample["extracted"]
        )

        results.append(
            result
        )

        # Save after every sample so progress
        # is not lost if the process stops.
        current_output = {
            "started_at": start_timestamp,
            "samples": results
        }

        RESULT_FILE.write_text(
            json.dumps(
                current_output,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

    # --------------------------------------------------------
    # Final summary
    # --------------------------------------------------------

    end_time = time.time()

    end_timestamp = (
        datetime.now()
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    total_duration = (
        end_time
        - start_time
    )

    successful = sum(
        1
        for result in results
        if result["status"] == "SUCCESS"
    )

    failed = sum(
        1
        for result in results
        if result["status"] != "SUCCESS"
    )

    print()
    print("=" * 70)
    print(
        "MILESTONE 7 TEST SUMMARY"
    )
    print("=" * 70)

    print(
        f"Start time: {start_timestamp}"
    )

    print(
        f"End time: {end_timestamp}"
    )

    print(
        f"Samples tested: {len(results)}"
    )

    print(
        f"Successful: {successful}"
    )

    print(
        f"Failed: {failed}"
    )

    print(
        f"Total duration: "
        f"{total_duration:.2f} seconds"
    )

    print()

    for result in results:

        print(
            f"Sample {result['sample_number']:02d} | "
            f"Pages: {result['page_count']} | "
            f"Status: {result['status']} | "
            f"Time: "
            f"{result['duration_seconds']:.2f}s"
        )

    print()

    print(
        f"Detailed results saved to:"
    )

    print(
        RESULT_FILE
    )

    print()

    if failed == 0:

        print(
            "ALL AVAILABLE SAMPLES "
            "COMPLETED SUCCESSFULLY."
        )

    else:

        print(
            "SOME SAMPLES FAILED."
        )

        print(
            "Review milestone7_accuracy_results.json."
        )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()