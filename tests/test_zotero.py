import unittest
from unittest.mock import Mock, patch
from io import BytesIO
import json

from zotero_word_citations.zotero import (
    ZoteroClient,
    add_doi_to_local_zotero,
    add_doi_with_cli,
)


class InMemoryClient(ZoteroClient):
    def __init__(self, items):
        self.items = items
        self.library_id = "0"

    def _request_json(self, path, params=None):
        return self.items, {"Total-Results": str(len(self.items))}


class ZoteroSelectionTests(unittest.TestCase):
    def test_duplicate_item_keys_are_sorted_alphabetically(self):
        items = [
            {"key": "ZAKPLF5D", "data": {"DOI": "10.1/example"}},
            {"key": "9P7SXS5B", "data": {"DOI": "10.1/EXAMPLE"}},
        ]
        matches = InMemoryClient(items).search_exact_doi("10.1/example")
        self.assertEqual(["9P7SXS5B", "ZAKPLF5D"], [item["key"] for item in matches])

    @patch("zotero_word_citations.zotero.shutil.which", return_value=r"C:\zotero-cli.exe")
    @patch("zotero_word_citations.zotero.subprocess.run")
    def test_cli_import_success_is_classified(self, run, _which):
        run.return_value = Mock(
            returncode=0,
            stdout="Successfully added: **Paper**\nItem key: `ABC12345`",
            stderr="",
        )
        result = add_doi_with_cli("https://doi.org/10.1000/EXAMPLE")
        self.assertEqual("added", result.status)
        self.assertEqual("ABC12345", result.item_key)
        command = run.call_args.args[0]
        self.assertIn("--if-exists", command)
        self.assertIn("skip", command)

    @patch("zotero_word_citations.zotero.shutil.which", return_value=r"C:\zotero-cli.exe")
    @patch("zotero_word_citations.zotero.subprocess.run")
    def test_cli_import_metadata_not_found_is_classified(self, run, _which):
        run.return_value = Mock(
            returncode=0,
            stdout="DOI not found on CrossRef: 10.9999/nope",
            stderr="",
        )
        result = add_doi_with_cli("10.9999/nope")
        self.assertEqual("metadata-not-found", result.status)

    @patch("zotero_word_citations.zotero.build_opener")
    def test_local_connector_fallback_fetches_and_saves_without_attachment(self, build):
        class Response(BytesIO):
            def __init__(self, data=b"", status=200):
                super().__init__(data)
                self.status = status

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                self.close()

        crossref = Mock()
        crossref.open.return_value = Response(
            json.dumps(
                {
                    "message": {
                        "type": "posted-content",
                        "title": ["Example Preprint"],
                        "author": [{"given": "A", "family": "Author"}],
                        "published": {"date-parts": [[2023, 11, 1]]},
                        "URL": "https://doi.org/10.1000/example",
                    }
                }
            ).encode("utf-8")
        )
        connector = Mock()
        connector.open.return_value = Response(status=201)
        build.side_effect = [crossref, connector]

        result = add_doi_to_local_zotero("10.1000/example")

        self.assertEqual("added", result.status)
        request = connector.open.call_args.args[0]
        payload = json.loads(request.data)
        self.assertEqual("preprint", payload["items"][0]["itemType"])
        self.assertEqual([], payload["items"][0]["attachments"])


if __name__ == "__main__":
    unittest.main()
