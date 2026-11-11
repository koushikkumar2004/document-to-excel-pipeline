"""
Unit tests for document_excel_pipeline.
Tests parsing, validation, date/numeric conversion, and schema integrity
without requiring a live LM Studio connection.
"""

import sys
import unittest
from datetime import date
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pipeline


class TestTypeConversions(unittest.TestCase):
    def setUp(self):
        pipeline.TYPE_WARNINGS.clear()

    def test_convert_number_integers(self):
        self.assertEqual(pipeline.convert_number("100", "quantity", 2), 100)
        self.assertEqual(pipeline.convert_number(5, "quantity", 2), 5)
        self.assertEqual(pipeline.convert_number("", "quantity", 2), "")
        self.assertEqual(pipeline.convert_number(None, "quantity", 2), "")

    def test_convert_number_floats_and_currency(self):
        self.assertEqual(pipeline.convert_number("$93.50", "total_amount", 2), 93.5)
        self.assertEqual(pipeline.convert_number("Rs. 42,000.00", "rate", 2), 42000)
        self.assertEqual(pipeline.convert_number("₹ 7,000", "subtotal_amount", 2), 7000)
        self.assertEqual(pipeline.convert_number("INR 1500.25", "total_amount", 2), 1500.25)
        self.assertEqual(pipeline.convert_number("€123.45", "line_amount", 2), 123.45)
        self.assertEqual(pipeline.convert_number("£99.99", "line_amount", 2), 99.99)

    def test_convert_number_rate_units(self):
        self.assertEqual(pipeline.convert_number("$50.00/hour", "rate", 2), 50)
        self.assertEqual(pipeline.convert_number("$25.00 / hrs", "rate", 2), 25)
        self.assertEqual(pipeline.convert_number("15.50/kg", "rate", 2), 15.5)

    def test_convert_number_percentages(self):
        self.assertEqual(pipeline.convert_number("20%", "discount_amount", 2), 0.2)
        self.assertEqual(pipeline.convert_number("5.5%", "discount_amount", 2), 0.055)

    def test_convert_number_preserves_words_with_warning(self):
        raw = "Seven Thousand Nine – Hundred Fifty Only"
        converted = pipeline.convert_number(raw, "total_amount", 2)
        self.assertEqual(converted, raw)
        self.assertTrue(any(w["value"] == raw for w in pipeline.TYPE_WARNINGS))

    def test_convert_number_strips_label_prefix(self):
        # LLM sometimes includes the field label: "Balance Due: $495.00"
        self.assertEqual(pipeline.convert_number("Balance Due: $495.00", "total_amount", 2), 495.0)
        self.assertEqual(pipeline.convert_number("Subtotal: Rs. 42,000.00", "total_amount", 2), 42000.0)


    def test_convert_date_standard_formats(self):
        self.assertEqual(pipeline.convert_date("2026-09-08", "document_date", 2), date(2026, 9, 8))
        self.assertEqual(pipeline.convert_date("January 25, 2016", "document_date", 2), date(2016, 1, 25))
        self.assertEqual(pipeline.convert_date("Jan 25, 2016", "document_date", 2), date(2016, 1, 25))
        self.assertEqual(pipeline.convert_date("5/6/03", "document_date", 2), date(2003, 5, 6))
        self.assertEqual(pipeline.convert_date("08 Sep 2026", "document_date", 2), date(2026, 9, 8))
        self.assertEqual(pipeline.convert_date("25/03/2026", "document_date", 2), date(2026, 3, 25))

    def test_convert_date_ordinal_formats(self):
        self.assertEqual(pipeline.convert_date("22th May, 2024", "document_date", 2), date(2024, 5, 22))
        self.assertEqual(pipeline.convert_date("25th May,2024", "document_date", 2), date(2024, 5, 25))
        self.assertEqual(pipeline.convert_date("1st June 2025", "document_date", 2), date(2025, 6, 1))
        self.assertEqual(pipeline.convert_date("2nd July 2025", "document_date", 2), date(2025, 7, 2))
        self.assertEqual(pipeline.convert_date("3rd August 2025", "document_date", 2), date(2025, 8, 3))

    def test_convert_date_preserves_invalid_with_warning(self):
        self.assertEqual(pipeline.convert_date("not-a-date", "document_date", 2), "not-a-date")
        self.assertTrue(any(w["value"] == "not-a-date" for w in pipeline.TYPE_WARNINGS))


class TestExtractionQuality(unittest.TestCase):
    def test_thousands_separator_not_flagged_as_multiple_values(self):
        batch = [{"page": 1, "text": "Item 1 Rate Rs. 42,000.00 Total Rs. 49,560.00"}]
        parsed = {
            "records": [
                {
                    "page": 1,
                    "header": {},
                    "line_items": [
                        {
                            "item_description": "Item 1",
                            "rate": "Rs. 42,000.00",
                            "line_amount": "Rs. 49,560.00",
                        }
                    ],
                }
            ]
        }
        issues = pipeline.extraction_quality_issues(parsed, batch)
        # Should NOT contain multiple comma-separated values issue
        self.assertFalse(any("multiple comma-separated values" in issue for issue in issues))

    def test_true_multiple_values_are_flagged(self):
        batch = [{"page": 1, "text": "Item: $14.95, $19.95, $15.99"}]
        parsed = {
            "records": [
                {
                    "page": 1,
                    "header": {},
                    "line_items": [
                        {
                            "item_description": "Merged item",
                            "rate": "$14.95, $19.95, $15.99",
                            "line_amount": "$50.89",
                        }
                    ],
                }
            ]
        }
        issues = pipeline.extraction_quality_issues(parsed, batch)
        self.assertTrue(any("multiple comma-separated values" in issue for issue in issues))

    def test_has_usable_extraction(self):
        empty = {"records": [{"page": 1, "header": {}, "line_items": []}]}
        self.assertFalse(pipeline.has_usable_extraction(empty))

        with_header = {"records": [{"page": 1, "header": {"document_number": "123"}, "line_items": []}]}
        self.assertTrue(pipeline.has_usable_extraction(with_header))

        with_item = {"records": [{"page": 1, "header": {}, "line_items": [{"item_code": "SKU-1"}]}]}
        self.assertTrue(pipeline.has_usable_extraction(with_item))

    def test_total_amount_in_words_is_flagged(self):
        batch = [{"page": 1, "text": "Total taxable value 2,90,000.00\nTwo lakh ninety thousand rupees only"}]
        parsed = {
            "records": [
                {
                    "page": 1,
                    "header": {"total_amount": "Two lakh ninety thousand rupees only"},
                    "line_items": [{"item_description": "Maize", "line_amount": "2,90,000.00"}],
                }
            ]
        }
        issues = pipeline.extraction_quality_issues(parsed, batch)
        self.assertTrue(any("extracted as words" in issue for issue in issues))

    def test_calculated_total_not_in_text_is_flagged(self):
        batch = [{"page": 1, "text": "Total before Tax: 28000\nTax: 5040\nTotal Amount after Tax: 33040"}]
        parsed = {
            "records": [
                {
                    "page": 1,
                    "header": {"total_amount": "38080"},
                    "line_items": [{"item_description": "Item", "line_amount": "33040"}],
                }
            ]
        }
        issues = pipeline.extraction_quality_issues(parsed, batch)
        self.assertTrue(any("has no matching numeric value in source text" in issue for issue in issues))

    def test_missing_price_rows_is_flagged(self):
        batch = [{"page": 1, "text": "Item 1 Description $50.00\nItem 2 Description $25.00\nItem 3 Description $320.00"}]
        parsed = {
            "records": [
                {
                    "page": 1,
                    "header": {},
                    "line_items": [
                        {"item_description": "Item 1", "rate": "50.00", "line_amount": "50.00"},
                        {"item_description": "Item 2", "rate": "25.00", "line_amount": "25.00"},
                    ],
                }
            ]
        }
        issues = pipeline.extraction_quality_issues(parsed, batch)
        self.assertTrue(any("price-bearing lines" in issue for issue in issues))


class TestJSONParsingAndFallbacks(unittest.TestCase):
    def test_parse_response_markdown_fences(self):
        raw = {
            "choices": [
                {
                    "message": {
                        "content": "```json\n{\"records\": [{\"page\": 1, \"header\": {\"document_number\": \"INV-1\"}, \"line_items\": []}], \"column_types\": {}}\n```"
                    }
                }
            ]
        }
        parsed = pipeline.parse_response(raw)
        self.assertIn("records", parsed)
        self.assertEqual(parsed["records"][0]["header"]["document_number"], "INV-1")

    def test_parse_response_think_tags(self):
        raw = {
            "choices": [
                {
                    "message": {
                        "content": "<think>Analyzing document...</think>{\"records\": [{\"page\": 1, \"header\": {\"document_number\": \"INV-99\"}, \"line_items\": []}], \"column_types\": {}}"
                    }
                }
            ]
        }
        parsed = pipeline.parse_response(raw)
        self.assertEqual(parsed["records"][0]["header"]["document_number"], "INV-99")

    def test_parse_response_recovers_header_from_column_types(self):
        raw = {
            "choices": [
                {
                    "message": {
                        "content": "{\"records\": [{\"page\": 1, \"header\": {}, \"line_items\": []}], \"column_types\": {\"document_number\": \"0001\", \"document_date\": \"5/6/03\", \"total_amount\": \"247.50\"}}"
                    }
                }
            ]
        }
        parsed = pipeline.parse_response(raw)
        header = parsed["records"][0]["header"]
        self.assertEqual(header.get("document_number"), "0001")
        self.assertEqual(header.get("document_date"), "5/6/03")
        self.assertEqual(header.get("total_amount"), "247.50")


class TestFlatteningAndSchema(unittest.TestCase):
    def test_flatten_fixed_preserves_columns_and_blanks(self):
        records = [
            {
                "page": 1,
                "header": {
                    "document_number": "INV-100",
                    "total_amount": "500",
                },
                "line_items": [
                    {
                        "item_description": "Widget A",
                        "quantity": "2",
                        "rate": "100",
                        "line_amount": "200",
                    },
                    {
                        "item_description": "Widget B",
                        "quantity": "3",
                        "rate": "100",
                        "line_amount": "300",
                    },
                ],
            }
        ]
        rows = pipeline.flatten_fixed(records)
        self.assertEqual(len(rows), 2)
        # Check that header fields propagated to all rows
        for row in rows:
            self.assertEqual(row["document_number"], "INV-100")
            self.assertEqual(row["total_amount"], "500")
            self.assertEqual(row["source_page"], 1)
            # Blank schema columns should be empty strings
            self.assertEqual(row["buyer_tax_id"], "")
            self.assertEqual(row["extra_1"], "")
            self.assertEqual(row["extra_10"], "")

        # Check line item distinction
        self.assertEqual(rows[0]["item_description"], "Widget A")
        self.assertEqual(rows[1]["item_description"], "Widget B")


if __name__ == "__main__":
    unittest.main()
