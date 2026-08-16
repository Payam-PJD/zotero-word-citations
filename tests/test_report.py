import tempfile
import unittest
from pathlib import Path

from zotero_word_citations.api import (
    DoiOutcome,
    PlaceholderOutcome,
    write_doi_report,
)
from zotero_word_citations.doi import DoiGroup


class DoiReportTests(unittest.TestCase):
    def test_report_marks_metadata_and_placeholder_statuses(self):
        group = DoiGroup(
            4,
            21,
            "[10.1000/example]",
            ("10.1000/example",),
            format="square",
        )
        outcome = DoiOutcome(
            doi="10.1000/example",
            status="metadata-fetched-and-added",
            title="Example Paper",
            selected_key="ABC12345",
        )
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "dois.txt"
            write_doi_report(
                path,
                input_path=Path("input.docx"),
                output_path=Path("output.docx"),
                groups=[group],
                outcomes={outcome.doi: outcome},
                placeholder_outcomes=[
                    PlaceholderOutcome(group, "cited", "Active field inserted.")
                ],
            )
            report = path.read_text(encoding="utf-8")
        self.assertIn("[METADATA-FETCHED-AND-ADDED] 10.1000/example", report)
        self.assertIn("Selected Zotero key: ABC12345", report)
        self.assertIn("[CITED] main text #1", report)


if __name__ == "__main__":
    unittest.main()
