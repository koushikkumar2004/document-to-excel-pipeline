from pathlib import Path


INPUT_FILE = Path("test_files/sample_invoice12_extracted.txt")


def split_pages(text):
    parts = text.split("--- Page ")

    pages = []

    for part in parts:
        if not part.strip():
            continue

        lines = part.split("\n", 1)

        page_number = lines[0].replace("---", "").strip()

        page_text = ""
        if len(lines) > 1:
            page_text = lines[1]

        pages.append({
            "page": int(page_number),
            "text": page_text.strip()
        })

    return pages


def create_batches(pages, batch_size=2):
    batches = []

    for i in range(0, len(pages), batch_size):
        batch = pages[i:i + batch_size]
        batches.append(batch)

    return batches


def main():
    text = INPUT_FILE.read_text(
        encoding="utf-8",
        errors="replace"
    )

    pages = split_pages(text)

    print(f"Total pages: {len(pages)}")

    batches = create_batches(
        pages,
        batch_size=2
    )

    print(f"Total batches: {len(batches)}")

    for number, batch in enumerate(batches, start=1):
        page_numbers = [page["page"] for page in batch]

        print(
            f"Batch {number}: "
            f"pages {page_numbers}"
        )

        print("-" * 40)


if __name__ == "__main__":
    main()