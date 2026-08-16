"""Print a structural inventory of Zotero Word citation fields in a DOCX."""

from __future__ import annotations

import gc
import json
from pathlib import Path
import sys

import win32com.client

from zotero_word_citations.word import _iter_story_ranges


def main() -> None:
    path = Path(sys.argv[1]).resolve()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(
            str(path), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        )
        visible_text = str(document.Content.Text)
        print(f"visible_text={visible_text!r}")
        fields = []
        for story_type, story_index, story in _iter_story_ranges(document):
            for index in range(1, int(story.Fields.Count) + 1):
                field = story.Fields(index)
                code = str(field.Code.Text)
                if not code.startswith(" ADDIN ZOTERO_ITEM CSL_CITATION "):
                    continue
                payload = json.loads(code[code.index("{") : code.rindex("}") + 1])
                fields.append(
                    {
                        "story": [story_type, story_index],
                        "type": int(field.Type),
                        "result": str(field.Result.Text),
                        "citationID": payload["citationID"],
                        "keys": [
                            item["uris"][0].rstrip("/").rsplit("/", 1)[-1]
                            for item in payload["citationItems"]
                        ],
                        "ids": [item["id"] for item in payload["citationItems"]],
                    }
                )
        print(json.dumps(fields, indent=2))
    finally:
        if document is not None:
            document.Close(SaveChanges=0)
        try:
            word.Quit()
        finally:
            document = None
            word = None
            gc.collect()


if __name__ == "__main__":
    main()
