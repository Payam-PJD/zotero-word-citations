import json
import unittest

from zotero_word_citations.word import (
    _existing_field_ranges,
    _field_code,
    _is_zotero_field,
    _merge_ranges,
    _overlaps_existing_field,
    _overlaps_existing_zotero_field,
    _unfielded_story_ranges,
)
from zotero_word_citations.zotero import ZoteroRecord


class WordPayloadTests(unittest.TestCase):
    def test_merged_citation_contains_multiple_items(self):
        records = [
            ZoteroRecord("AAA", "10.1/a", "A", "http://zotero/items/AAA", {"id": "old", "title": "A"}, 1),
            ZoteroRecord("BBB", "10.1/b", "B", "http://zotero/items/BBB", {"id": "old", "title": "B"}, 2),
        ]
        code = _field_code(records, "citation", "3,4")
        self.assertTrue(code.startswith(" ADDIN ZOTERO_ITEM CSL_CITATION "))
        payload = json.loads(code[code.index("{") : code.rindex("}") + 1])
        self.assertEqual([1, 2], [item["id"] for item in payload["citationItems"]])
        self.assertEqual("3,4", payload["properties"]["plainCitation"])

    def test_field_envelopes_are_merged_and_protected(self):
        class Range:
            def __init__(self, start, end, text=""):
                self.Start = start
                self.End = end
                self.Text = text

        class Field:
            def __init__(self, code_start, result_end, code=" PAGE "):
                self.Code = Range(code_start, result_end - 1, code)
                self.Result = Range(code_start + 1, result_end)

        class Fields:
            def __init__(self, values):
                self.values = values
                self.Count = len(values)

            def __call__(self, index):
                return self.values[index - 1]

        class Story:
            pass

        story = Story()
        story.Start = 0
        story.End = 100
        story.Fields = Fields(
            [
                Field(11, 20, " ADDIN ZOTERO_ITEM CSL_CITATION {} "),
                Field(20, 30, " ADDIN ZOTERO_BIBL {} "),
                Field(51, 60),
            ]
        )

        class DuplicateRange:
            def __init__(self):
                self.Start = story.Start
                self.End = story.End

            def SetRange(self, start, end):
                self.Start = start
                self.End = end

        class DuplicateDescriptor:
            def __get__(self, instance, owner):
                return DuplicateRange()

        Story.Duplicate = DuplicateDescriptor()

        # Complete field envelopes expand one position around Code/Result.
        self.assertEqual([(10, 31), (50, 61)], _existing_field_ranges(story))
        self.assertTrue(_overlaps_existing_field(story, 15, 16))
        self.assertTrue(_overlaps_existing_field(story, 60, 62))
        self.assertFalse(_overlaps_existing_field(story, 31, 50))
        self.assertTrue(_overlaps_existing_zotero_field(story, 15, 16))
        self.assertFalse(_overlaps_existing_zotero_field(story, 55, 56))
        self.assertTrue(_is_zotero_field(story.Fields(1)))
        self.assertTrue(_is_zotero_field(story.Fields(2)))
        self.assertFalse(_is_zotero_field(story.Fields(3)))
        self.assertEqual(
            [(0, 10), (31, 50), (61, 100)],
            [(item.Start, item.End) for item in _unfielded_story_ranges(story)],
        )

    def test_merge_ranges_handles_nested_and_adjacent_fields(self):
        self.assertEqual(
            [(1, 12), (20, 25)],
            _merge_ranges([(5, 10), (1, 6), (10, 12), (20, 25), (22, 24)]),
        )


if __name__ == "__main__":
    unittest.main()
