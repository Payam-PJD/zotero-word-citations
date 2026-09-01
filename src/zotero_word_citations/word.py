from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, replace
import gc
import html
import json
from pathlib import Path
import secrets
import shutil
import string
import os
import tempfile
from typing import Iterable, Iterator, Mapping, Sequence
from uuid import uuid4

from .doi import DoiGroup, find_doi_like_values, parse_doi_groups
from .zotero import ZoteroRecord


WD_COLLAPSE_END = 0
WD_DO_NOT_SAVE_CHANGES = 0
WD_FIELD_ADDIN = 81
WD_FIELD_QUOTE = 35
WD_FORMAT_DOCUMENT_DEFAULT = 16
WD_REVISIONS_VIEW_FINAL = 0
MSO_PROPERTY_TYPE_STRING = 4
POSITION_MAP_CHUNK = 256
CSL_CITATION_SCHEMA = (
    "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
)
DEFAULT_STYLE_ID = "http://www.zotero.org/styles/american-medical-association"

# Word story types that can contain normal document prose. Comments and the
# generated footnote/endnote separator stories are intentionally omitted.
SCANNED_STORY_TYPES = (1, 2, 3, 5, 6, 7, 8, 9, 10, 11)
WORD_STORY_NAMES = {
    1: "main text",
    2: "footnotes",
    3: "endnotes",
    5: "text frame",
    6: "even-page header",
    7: "primary header",
    8: "even-page footer",
    9: "primary footer",
    10: "first-page header",
    11: "first-page footer",
}


class WordError(RuntimeError):
    """Raised when Microsoft Word cannot create the requested fields."""


@dataclass(frozen=True, slots=True)
class InsertedCitation:
    source: str
    dois: tuple[str, ...]
    item_keys: tuple[str, ...]
    citation_id: str
    provisional_text: str


class _PositionMap:
    """Map displayed-text indices in a Word range to Word positions.

    Word positions include hidden tracked deletions and field instructions,
    while ``Range.Text`` follows the current revision view. A Python string
    index therefore cannot safely be added to ``Range.Start``. This map uses
    bounded COM probes and refuses inconsistent/non-additive ranges.
    """

    def __init__(self, com_range, chunk_size: int = POSITION_MAP_CHUNK) -> None:
        self._source = com_range.Duplicate
        self._start = int(com_range.Start)
        self._end = int(com_range.End)
        self.text = str(com_range.Text or "")
        self._marks = [self._start]
        self._cumulative = [0]

        total = 0
        position = self._start
        probe = com_range.Duplicate
        while position < self._end:
            next_position = min(position + max(1, chunk_size), self._end)
            probe.SetRange(position, next_position)
            total += len(str(probe.Text or ""))
            self._marks.append(next_position)
            self._cumulative.append(total)
            position = next_position
        self.consistent = total == len(self.text)

    def _displayed_before(self, position: int) -> int:
        low, high = 0, len(self._marks) - 1
        while low < high:
            middle = (low + high + 1) // 2
            if self._marks[middle] <= position:
                low = middle
            else:
                high = middle - 1
        total = self._cumulative[low]
        if self._marks[low] == position:
            return total
        probe = self._source.Duplicate
        probe.SetRange(self._marks[low], position)
        return total + len(str(probe.Text or ""))

    def start_position(self, index: int) -> int:
        """Return the Word position of displayed character ``index``."""

        if not 0 <= index <= len(self.text):
            raise IndexError(index)
        low, high = self._start, self._end
        while low < high:
            middle = (low + high + 1) // 2
            if self._displayed_before(middle) <= index:
                low = middle
            else:
                high = middle - 1
        return low

    def end_position(self, index: int) -> int:
        """Return the first Word position after ``index`` displayed chars."""

        if not 0 <= index <= len(self.text):
            raise IndexError(index)
        low, high = self._start, self._end
        while low < high:
            middle = (low + high) // 2
            if self._displayed_before(middle) >= index:
                high = middle
            else:
                low = middle + 1
        return low


def _force_final_view(document) -> dict[str, object]:
    """Show final text deterministically and return state for restoration."""

    try:
        view = document.ActiveWindow.View
    except Exception as exc:
        raise WordError("Microsoft Word did not expose a document view.") from exc

    saved: dict[str, object] = {"view": view}
    for name in ("ShowRevisionsAndComments", "RevisionsView"):
        try:
            saved[name] = getattr(view, name)
        except Exception:
            pass
    try:
        saved["Markup"] = view.RevisionsFilter.Markup
    except Exception:
        pass

    try:
        view.RevisionsView = WD_REVISIONS_VIEW_FINAL
        view.ShowRevisionsAndComments = False
    except Exception as exc:
        raise WordError(
            "Microsoft Word could not be pinned to Final revision view; "
            "conversion was stopped to avoid citing deleted text."
        ) from exc
    return saved


def _restore_view(saved: dict[str, object] | None) -> None:
    if not saved:
        return
    view = saved.get("view")
    if view is None:
        return
    if "Markup" in saved:
        try:
            view.RevisionsFilter.Markup = saved["Markup"]
        except Exception:
            pass
    for name in ("RevisionsView", "ShowRevisionsAndComments"):
        if name in saved:
            try:
                setattr(view, name, saved[name])
            except Exception:
                pass


def _iter_story_ranges(document) -> Iterator[tuple[int, int, object]]:
    """Yield each editable Word story range with a stable type/index key."""

    for story_type in SCANNED_STORY_TYPES:
        try:
            story = document.StoryRanges(story_type)
        except Exception:
            continue
        story_index = 0
        while story is not None:
            yield story_type, story_index, story
            try:
                story = story.NextStoryRange
            except Exception:
                story = None
            story_index += 1


def _get_story_range(document, story_type: int, story_index: int):
    try:
        story = document.StoryRanges(story_type)
    except Exception as exc:
        raise WordError(
            f"Word story {story_type}/{story_index} no longer exists."
        ) from exc
    for _ in range(story_index):
        story = story.NextStoryRange
        if story is None:
            raise WordError(
                f"Word story {story_type}/{story_index} no longer exists."
            )
    return story


def _field_envelope(field) -> tuple[int, int]:
    """Return the absolute Word range occupied by a complete field.

    Word's ``Field.Code`` and ``Field.Result`` omit the begin/end field
    characters themselves. Expanding by one character on each side protects
    the complete field, including its hidden instruction text and visible
    result.
    """

    return int(field.Code.Start) - 1, int(field.Result.End) + 1


def _merge_ranges(ranges: Iterable[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[tuple[int, int]] = []
    for start, end in sorted(ranges):
        if end <= start:
            continue
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _existing_field_ranges(story) -> list[tuple[int, int]]:
    """Return absolute envelopes for every existing Word field in a story.

    These envelopes split literal story text from hidden field instructions.
    Zotero envelopes are always protected; a separate narrow path can replace
    a non-Zotero field whose complete displayed result is a DOI placeholder.
    """

    excluded: list[tuple[int, int]] = []
    try:
        field_count = int(story.Fields.Count)
    except Exception:
        return excluded
    for index in range(1, field_count + 1):
        field = story.Fields(index)
        try:
            excluded.append(_field_envelope(field))
        except Exception:
            continue
    return _merge_ranges(excluded)


def _is_zotero_field(field) -> bool:
    try:
        code = str(field.Code.Text).lstrip().casefold()
    except Exception:
        return False
    return code.startswith("addin zotero_")


def _existing_zotero_field_ranges(story) -> list[tuple[int, int]]:
    excluded: list[tuple[int, int]] = []
    try:
        field_count = int(story.Fields.Count)
    except Exception:
        return excluded
    for index in range(1, field_count + 1):
        field = story.Fields(index)
        try:
            if _is_zotero_field(field):
                excluded.append(_field_envelope(field))
        except Exception:
            continue
    return _merge_ranges(excluded)


def _unfielded_story_ranges(story) -> Iterator[object]:
    """Yield literal-text ranges outside every active Word field.

    ``story.Text`` hides field instructions while Word's character coordinates
    still count them. Splitting on complete field envelopes avoids that offset
    mismatch and makes every parsed DOI position a real Word position.
    """

    story_start = int(story.Start)
    story_end = int(story.End)
    cursor = story_start
    for field_start, field_end in _existing_field_ranges(story):
        field_start = max(story_start, field_start)
        field_end = min(story_end, field_end)
        if cursor < field_start:
            safe_range = story.Duplicate
            safe_range.SetRange(cursor, field_start)
            yield safe_range
        cursor = max(cursor, field_end)
    if cursor < story_end:
        safe_range = story.Duplicate
        safe_range.SetRange(cursor, story_end)
        yield safe_range


def _overlaps_existing_field(story, start: int, end: int) -> bool:
    return any(
        start < field_end and end > field_start
        for field_start, field_end in _existing_field_ranges(story)
    )


def _overlaps_existing_zotero_field(story, start: int, end: int) -> bool:
    return any(
        start < field_end and end > field_start
        for field_start, field_end in _existing_zotero_field_ranges(story)
    )


def _mapped_groups_in_range(
    safe_range,
    *,
    story_start: int,
    story_type: int,
    story_index: int,
    formats: Iterable[str] | None,
    warnings: list[str],
) -> list[DoiGroup]:
    mapper = _PositionMap(safe_range)
    if not mapper.text:
        return []
    if not mapper.consistent:
        raise WordError(
            "Word reported inconsistent displayed-text lengths in story "
            f"{story_type}/{story_index}; refusing to guess DOI offsets."
        )

    mapped: list[DoiGroup] = []
    verifier = safe_range.Duplicate
    for group in parse_doi_groups(
        mapper.text,
        formats=formats,
        story_type=story_type,
        story_index=story_index,
    ):
        begin = mapper.start_position(group.start)
        finish = mapper.end_position(group.end)
        verifier.SetRange(begin, finish)
        actual = str(verifier.Text or "")
        if actual != group.source or finish - begin != len(group.source):
            warnings.append(
                "Skipped a DOI placeholder containing hidden/interleaved Word "
                f"content in story {story_type}/{story_index}: {group.source}"
            )
            continue
        mapped.append(
            replace(
                group,
                start=begin - story_start,
                end=finish - story_start,
            )
        )
    return mapped


def scan_word_document(
    input_path: Path,
    formats: Iterable[str] | None = None,
    warnings: list[str] | None = None,
) -> list[DoiGroup]:
    """Read all normal Word stories and locate DOI placeholders safely.

    Existing Zotero fields are excluded in full, so neither their hidden JSON
    nor displayed results can be mistaken for new placeholders. Non-Zotero
    fields are accepted only when the complete result is a DOI placeholder.
    Returned offsets use Word's true story coordinates.
    """

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise WordError("pywin32 is required to read Microsoft Word files.") from exc

    source = input_path.expanduser().resolve()
    if not source.is_file():
        raise WordError(f"Input document does not exist: {source}")
    if source.suffix.casefold() != ".docx":
        raise WordError("The input must use the .docx extension.")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    groups: list[DoiGroup] = []
    scan_warnings = warnings if warnings is not None else []
    view_state: dict[str, object] | None = None
    try:
        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        view_state = _force_final_view(document)
        for story_type, story_index, story in _iter_story_ranges(document):
            story_start = int(story.Start)
            for safe_range in _unfielded_story_ranges(story):
                groups.extend(
                    _mapped_groups_in_range(
                        safe_range,
                        story_start=story_start,
                        story_type=story_type,
                        story_index=story_index,
                        formats=formats,
                        warnings=scan_warnings,
                    )
                )

            # A pasted DOI URL can be an active Word HYPERLINK field. We never
            # inspect Zotero fields, but a non-Zotero field is replaceable when
            # its *entire displayed result* is exactly one DOI placeholder.
            # The replacement coordinates cover the complete outer field, so
            # no hidden HYPERLINK instruction is left behind.
            try:
                field_count = int(story.Fields.Count)
            except Exception:
                field_count = 0
            for field_index in range(1, field_count + 1):
                field = story.Fields(field_index)
                if _is_zotero_field(field):
                    continue
                try:
                    if int(field.Result.Fields.Count):
                        continue
                except Exception:
                    pass
                try:
                    result_text = str(field.Result.Text)
                    field_start, field_end = _field_envelope(field)
                except Exception:
                    continue
                parsed = parse_doi_groups(
                    result_text,
                    formats=formats,
                    story_type=story_type,
                    story_index=story_index,
                )
                for group in parsed:
                    if group.start == 0 and group.end == len(result_text):
                        groups.append(
                            replace(
                                group,
                                start=field_start - story_start,
                                end=field_end - story_start,
                            )
                        )
    except Exception as exc:
        if isinstance(exc, (WordError, ValueError)):
            raise
        raise WordError(f"Microsoft Word scan failed: {exc}") from exc
    finally:
        if document is not None:
            try:
                _restore_view(view_state)
                document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        try:
            word.Quit()
        finally:
            document = None
            word = None
            gc.collect()
    return sorted(
        groups,
        key=lambda group: (group.story_type, group.story_index, group.start, group.end),
    )


def scan_word_doi_like_values(input_path: Path) -> tuple[str, ...]:
    """Return DOI-shaped visible text outside all Word fields in Final view."""

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise WordError("pywin32 is required to read Microsoft Word files.") from exc

    source = input_path.expanduser().resolve()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    view_state: dict[str, object] | None = None
    values: list[str] = []
    try:
        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        view_state = _force_final_view(document)
        for _, _, story in _iter_story_ranges(document):
            for safe_range in _unfielded_story_ranges(story):
                values.extend(find_doi_like_values(str(safe_range.Text or "")))
    except Exception as exc:
        if isinstance(exc, WordError):
            raise
        raise WordError(f"Microsoft Word reconciliation scan failed: {exc}") from exc
    finally:
        if document is not None:
            try:
                _restore_view(view_state)
                document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        try:
            word.Quit()
        finally:
            document = None
            word = None
            gc.collect()
    return tuple(values)


def _random_id(length: int = 8) -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(length))


def _citation_payload(
    records: Sequence[ZoteroRecord], citation_id: str, provisional_text: str
) -> dict:
    citation_items = []
    for record in records:
        item_data = deepcopy(record.item_data)
        item_data["id"] = record.field_id
        citation_items.append(
            {
                "id": record.field_id,
                "uris": [record.uri],
                "itemData": item_data,
            }
        )
    return {
        "citationID": citation_id,
        "properties": {
            "unsorted": False,
            "formattedCitation": f"\\super {provisional_text}\\nosupersub{{}}",
            "plainCitation": provisional_text,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": CSL_CITATION_SCHEMA,
    }


def _field_code(records: Sequence[ZoteroRecord], citation_id: str, text: str) -> str:
    payload = _citation_payload(records, citation_id, text)
    compact = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return f" ADDIN ZOTERO_ITEM CSL_CITATION {compact} "


def _get_custom_property(properties, name: str):
    try:
        return properties(name)
    except Exception:
        return None


def _read_document_data(document) -> str:
    properties = document.CustomDocumentProperties
    chunks: list[str] = []
    index = 1
    while True:
        prop = _get_custom_property(properties, f"ZOTERO_PREF_{index}")
        if prop is None:
            break
        chunks.append(str(prop.Value))
        index += 1
    return "".join(chunks)


def _write_document_data(document, value: str) -> None:
    properties = document.CustomDocumentProperties
    chunks = [value[index : index + 255] for index in range(0, len(value), 255)]
    for index, chunk in enumerate(chunks, start=1):
        name = f"ZOTERO_PREF_{index}"
        prop = _get_custom_property(properties, name)
        if prop is None:
            properties.Add(name, False, MSO_PROPERTY_TYPE_STRING, chunk)
        else:
            prop.Value = chunk

    index = len(chunks) + 1
    while True:
        prop = _get_custom_property(properties, f"ZOTERO_PREF_{index}")
        if prop is None:
            break
        prop.Delete()
        index += 1


def _ensure_document_data(document, zotero_version: str, style_id: str) -> bool:
    existing = _read_document_data(document)
    if existing:
        if 'name="fieldType" value="Field"' not in existing:
            raise WordError(
                "The document contains Zotero preferences for a non-Word field type."
            )
        return False

    version = html.escape(zotero_version if zotero_version != "unknown" else "9.0", quote=True)
    style = html.escape(style_id, quote=True)
    session = _random_id()
    document_data = (
        f'<data data-version="3" zotero-version="{version}">'
        f'<session id="{session}"/>'
        f'<style id="{style}" hasBibliography="1" bibliographyStyleHasBeenSet="0"/>'
        '<prefs><pref name="fieldType" value="Field"/></prefs></data>'
    )
    _write_document_data(document, document_data)
    return True


def _validate_persisted_fields(document, expected: Sequence[InsertedCitation]) -> None:
    """Verify that a saved/reopened document retained every new ADDIN field."""

    expected_by_id = {citation.citation_id: citation for citation in expected}
    found: set[str] = set()
    fields = []
    for _, _, story in _iter_story_ranges(document):
        for index in range(1, int(story.Fields.Count) + 1):
            fields.append(story.Fields(index))
    for field in fields:
        code = str(field.Code.Text)
        if not code.startswith(" ADDIN ZOTERO_ITEM CSL_CITATION "):
            continue
        json_start = code.find("{")
        json_end = code.rfind("}")
        if json_start < 0 or json_end <= json_start:
            continue
        try:
            payload = json.loads(code[json_start : json_end + 1])
        except json.JSONDecodeError:
            continue
        citation_id = payload.get("citationID")
        if citation_id not in expected_by_id:
            continue
        expected_citation = expected_by_id[citation_id]
        if int(field.Type) != WD_FIELD_ADDIN:
            raise WordError(
                f"Citation {citation_id} did not persist as a Word ADDIN field."
            )
        actual_keys = tuple(
            str(item.get("uris", [""])[0]).rstrip("/").rsplit("/", 1)[-1]
            for item in payload.get("citationItems", [])
        )
        if actual_keys != expected_citation.item_keys:
            raise WordError(
                f"Citation {citation_id} persisted with unexpected Zotero items."
            )
        found.add(citation_id)

    missing = set(expected_by_id) - found
    if missing:
        raise WordError(
            "Word did not persist all inserted Zotero fields: " + ", ".join(sorted(missing))
        )
    if not _read_document_data(document):
        raise WordError("The saved document is missing Zotero document preferences.")


def convert_word_document(
    input_path: Path,
    output_path: Path,
    groups: Sequence[DoiGroup],
    records_by_doi: Mapping[str, ZoteroRecord],
    zotero_version: str,
    style_id: str = DEFAULT_STYLE_ID,
    force: bool = False,
) -> tuple[list[InsertedCitation], bool]:
    """Replace DOI groups with real Word ADDIN fields in a new DOCX."""

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise WordError(
            "pywin32 is required. Install this package with `python -m pip install .`."
        ) from exc

    source = input_path.expanduser().resolve()
    destination = output_path.expanduser().absolute()
    if not source.is_file():
        raise WordError(f"Input document does not exist: {source}")
    if source.suffix.casefold() != ".docx" or destination.suffix.casefold() != ".docx":
        raise WordError("Both input and output must use the .docx extension.")
    if source == destination.resolve():
        raise WordError("The output must differ from the input; source files are preserved.")
    if destination.exists() and not force:
        raise WordError(f"Output already exists: {destination}. Use --force to replace it.")
    destination.parent.mkdir(parents=True, exist_ok=True)

    number_by_key: dict[str, int] = {}
    next_number = 1
    prepared: list[tuple[int, DoiGroup, list[ZoteroRecord], str, str]] = []
    for ordinal, group in enumerate(groups):
        group_records: list[ZoteroRecord] = []
        seen_keys: set[str] = set()
        for doi in group.dois:
            record = records_by_doi[doi]
            if record.key not in seen_keys:
                group_records.append(record)
                seen_keys.add(record.key)
            if record.key not in number_by_key:
                number_by_key[record.key] = next_number
                next_number += 1
        provisional = ",".join(str(number_by_key[record.key]) for record in group_records)
        prepared.append((ordinal, group, group_records, _random_id(), provisional))

    # Keep the working DOCX outside OneDrive-synced output folders. Some cloud
    # providers resurrect a hidden source path after an in-folder atomic move.
    temporary_dir = Path(tempfile.mkdtemp(prefix="zotero-word-cite-word-"))
    temporary = temporary_dir / f"{destination.stem}.{uuid4().hex}.tmp.docx"
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    inserted: list[InsertedCitation] = []
    created_preferences = False
    view_state: dict[str, object] | None = None
    original_track_revisions: bool | None = None
    try:
        document = word.Documents.Open(
            str(source), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        )
        view_state = _force_final_view(document)
        try:
            original_track_revisions = bool(document.TrackRevisions)
        except Exception:
            original_track_revisions = None
        document.SaveAs2(str(temporary), FileFormat=WD_FORMAT_DOCUMENT_DEFAULT)
        # Never record the placeholder replacement itself as tracked changes.
        # Existing revisions are preserved. Word may split an existing tracked
        # insertion around a new field, but its revision content is unchanged.
        document.TrackRevisions = False
        created_preferences = _ensure_document_data(document, zotero_version, style_id)

        inserted_by_ordinal: list[tuple[int, InsertedCitation]] = []
        # Work backwards within each story so earlier character offsets remain
        # valid. Story-relative ranges also support footnotes, text boxes, and
        # headers instead of assuming everything is in document.Content.
        story_keys = sorted(
            {(group.story_type, group.story_index) for group in groups}
        )
        for story_type, story_index in story_keys:
            story = _get_story_range(document, story_type, story_index)
            story_start = int(story.Start)
            story_prepared = [
                value
                for value in prepared
                if value[1].story_type == story_type
                and value[1].story_index == story_index
            ]
            for _, group, _, _, _ in story_prepared:
                absolute_start = story_start + group.start
                absolute_end = story_start + group.end
                if _overlaps_existing_zotero_field(
                    story, absolute_start, absolute_end
                ):
                    raise WordError(
                        "A requested DOI replacement overlaps an existing Zotero "
                        "field. Conversion was aborted without changing the source."
                    )
            for ordinal, group, records, citation_id, provisional in reversed(
                story_prepared
            ):
                target = story.Duplicate
                target.SetRange(
                    story_start + group.start,
                    story_start + group.end,
                )
                if target.Text != group.source:
                    raise WordError(
                        "Word text offsets changed before insertion; conversion was aborted."
                    )
                target.Text = ""
                target.SetRange(target.Start, target.Start)
                field = document.Fields.Add(
                    target,
                    Type=WD_FIELD_QUOTE,
                    Text="{Citation}",
                    PreserveFormatting=True,
                )
                field.Code.Text = _field_code(records, citation_id, provisional)
                field.Result.Text = provisional
                field.Result.Font.Superscript = True
                inserted_by_ordinal.append(
                    (
                        ordinal,
                        InsertedCitation(
                            source=group.source,
                            dois=group.dois,
                            item_keys=tuple(record.key for record in records),
                            citation_id=citation_id,
                            provisional_text=provisional,
                        ),
                    )
                )

        inserted = [
            citation
            for _, citation in sorted(inserted_by_ordinal, key=lambda pair: pair[0])
        ]

        # Release the final range/field COM proxies before closing Word.
        field = None
        target = None

        if original_track_revisions is not None:
            document.TrackRevisions = original_track_revisions
        _restore_view(view_state)
        view_state = None
        document.Save()
        document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        document = None
        original_track_revisions = None

        # A field changes from the temporary QUOTE type (35) to Word's ADDIN
        # type (81) only after save/reopen. Validate before committing output.
        document = word.Documents.Open(
            str(temporary), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        )
        view_state = _force_final_view(document)
        _validate_persisted_fields(document, inserted)
        _restore_view(view_state)
        view_state = None
        document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        document = None
        os.replace(temporary, destination)
    except Exception as exc:
        if document is not None:
            try:
                if original_track_revisions is not None:
                    document.TrackRevisions = original_track_revisions
                _restore_view(view_state)
                document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if temporary.exists():
            temporary.unlink()
        if isinstance(exc, WordError):
            raise
        raise WordError(f"Microsoft Word conversion failed: {exc}") from exc
    finally:
        try:
            word.Quit()
        finally:
            document = None
            word = None
            gc.collect()
            shutil.rmtree(temporary_dir, ignore_errors=True)

    return inserted, created_preferences
