import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from zotero_word_citations.api import _safe_console_print, cite_document
from zotero_word_citations.doi import DoiGroup
from zotero_word_citations.word import InsertedCitation
from zotero_word_citations.zotero import (
    DoiImportResult,
    DoiResolution,
    ZoteroRecord,
)


class CiteDocumentTests(unittest.TestCase):
    def test_progress_output_survives_legacy_windows_encoding(self):
        raw = io.BytesIO()
        stream = io.TextIOWrapper(raw, encoding="ascii")
        with patch("zotero_word_citations.api.sys.stdout", stream):
            _safe_console_print("bone′s")
            stream.flush()
        self.assertEqual(b"bone?s\r\n", raw.getvalue())

    @patch("zotero_word_citations.api.scan_word_doi_like_values", return_value=())
    @patch("zotero_word_citations.api.convert_word_document")
    @patch("zotero_word_citations.api.resolve_numeric_item_ids", return_value={})
    @patch("zotero_word_citations.api.discover_zotero_data_dir", return_value=None)
    @patch("zotero_word_citations.api.add_doi_with_cli")
    @patch("zotero_word_citations.api.add_doi_to_local_zotero")
    @patch("zotero_word_citations.api.write_doi_report")
    @patch("zotero_word_citations.api.scan_word_document")
    @patch("zotero_word_citations.api.ZoteroClient")
    def test_local_connector_is_primary_and_cli_is_not_called_on_success(
        self,
        client_type,
        scan,
        _report,
        local_add,
        cli_add,
        _data_dir,
        _numeric_ids,
        convert,
        _reconcile,
    ):
        doi = "10.1000/example"
        group = DoiGroup(0, 17, doi, (doi,), format="bare")
        scan.return_value = [group]
        record = ZoteroRecord(
            key="ABC12345",
            doi=doi,
            title="Example",
            uri="http://zotero.org/users/local/items/ABC12345",
            item_data={"id": 1, "type": "article-journal", "title": "Example"},
            item_id=1,
        )
        resolution = DoiResolution(doi, record.key, (record.key,))
        client = client_type.return_value
        client.version = "9.0.6"
        client.resolve_doi_if_present.side_effect = [None, (record, resolution)]
        local_add.return_value = DoiImportResult(
            doi, "added", "Saved through local Zotero."
        )
        convert.return_value = (
            [InsertedCitation(doi, (doi,), (record.key,), "citation", "1")],
            False,
        )

        with tempfile.TemporaryDirectory() as temp_name:
            source = Path(temp_name) / "input.docx"
            source.write_bytes(b"placeholder")
            result = cite_document(source, verbose=False)

        client.wait_until_available.assert_called_once()
        local_add.assert_called_once()
        cli_add.assert_not_called()
        self.assertEqual(1, len(result.inserted_citations))


if __name__ == "__main__":
    unittest.main()
