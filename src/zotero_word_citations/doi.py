from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable, Sequence


DOI_FORMATS = (
    "bare",
    "url",
    "doi-prefix",
    "round",
    "square",
    "curly",
    "bibtex",
)

# Crossref's practical DOI grammar is deliberately broad. The detector first
# finds every DOI start, then applies punctuation and grouping heuristics. This
# catches old DOI suffixes containing parentheses while still handling adjacent
# placeholders such as ``10.1/a;10.2/b``.
_DOI_START_RE = re.compile(
    r"(?<![A-Za-z0-9])"
    r"(?:(?P<url>https?://(?:dx\.)?doi\.org/)|(?P<label>doi\s*:\s*))?"
    r"(?P<doi>10\.\d{4,9}/)",
    re.IGNORECASE,
)
_SUFFIX_RE = re.compile(r"[-._;()/:A-Z0-9<>%+@]+", re.IGNORECASE)
_BIBTEX_PREFIX_RE = re.compile(
    r"(?i)(?:\\?(?:cite|citep|citet|autocite|parencite|textcite|footcite)\*?\s*)$"
)
_SEPARATOR_RE = re.compile(r"(?:\s*[,;]\s*|\s*-\s*)+")
_TRAILING_PUNCTUATION = ".,;:!?"
_BRACKET_PAIRS = {"(": ")", "[": "]", "{": "}"}
_FORMAT_FOR_OPEN = {"(": "round", "[": "square", "{": "curly"}


@dataclass(frozen=True, slots=True)
class DoiToken:
    """One DOI-like token and its offsets in a Word story's text."""

    start: int
    end: int
    source: str
    doi: str
    format: str


@dataclass(frozen=True, slots=True)
class DoiGroup:
    """A DOI placeholder group and its offsets in a Word story's text."""

    start: int
    end: int
    source: str
    dois: tuple[str, ...]
    format: str = "round"
    story_type: int = 1
    story_index: int = 0


def normalize_doi(doi: str) -> str:
    """Return the canonical comparison form used for exact DOI matching."""

    value = doi.strip()
    value = re.sub(r"(?i)^https?://(?:dx\.)?doi\.org/", "", value)
    value = re.sub(r"(?i)^doi\s*:\s*", "", value)
    return value.casefold()


def _trim_suffix(suffix: str) -> str:
    """Remove sentence/wrapper punctuation without breaking balanced DOI parens."""

    suffix = suffix.rstrip(_TRAILING_PUNCTUATION)
    while suffix.endswith(")") and suffix.count(")") > suffix.count("("):
        suffix = suffix[:-1].rstrip(_TRAILING_PUNCTUATION)
    return suffix


def find_doi_tokens(text: str) -> list[DoiToken]:
    """Find DOI-like strings in bare, ``doi:``, and URL forms.

    A syntactically plausible DOI is returned even when it has no metadata in
    Zotero or Crossref. That distinction is made later and recorded in the run
    report.
    """

    starts = list(_DOI_START_RE.finditer(text))
    tokens: list[DoiToken] = []
    for index, match in enumerate(starts):
        suffix_start = match.end("doi")
        suffix_match = _SUFFIX_RE.match(text, suffix_start)
        if suffix_match is None:
            continue
        raw_end = suffix_match.end()

        # A following DOI start caps the current candidate. This handles a
        # hyphen/semicolon delimiter even when no spaces are present.
        if index + 1 < len(starts):
            raw_end = min(raw_end, starts[index + 1].start())
        suffix = text[suffix_start:raw_end]
        suffix = suffix.rstrip(" \t\r\n,;-")
        suffix = _trim_suffix(suffix)
        if not suffix:
            continue

        end = suffix_start + len(suffix)
        source = text[match.start():end]
        kind = (
            "url"
            if match.group("url")
            else "doi-prefix"
            if match.group("label")
            else "bare"
        )
        tokens.append(
            DoiToken(
                start=match.start(),
                end=end,
                source=source,
                doi=normalize_doi(text[match.start("doi"):end]),
                format=kind,
            )
        )
    return tokens


def _bracket_pairs(text: str) -> list[tuple[int, int, str]]:
    stack: list[tuple[str, int]] = []
    pairs: list[tuple[int, int, str]] = []
    closing = {value: key for key, value in _BRACKET_PAIRS.items()}
    for index, char in enumerate(text):
        if char in _BRACKET_PAIRS:
            stack.append((char, index))
        elif char in closing:
            expected = closing[char]
            for stack_index in range(len(stack) - 1, -1, -1):
                if stack[stack_index][0] == expected:
                    open_char, start = stack[stack_index]
                    del stack[stack_index:]
                    pairs.append((start, index + 1, open_char))
                    break
    return pairs


def _tokens_in_span(
    tokens: Sequence[DoiToken], start: int, end: int
) -> list[DoiToken]:
    return [token for token in tokens if token.start >= start and token.end <= end]


def _only_dois_and_separators(
    text: str, content_start: int, content_end: int, tokens: Sequence[DoiToken]
) -> bool:
    if not tokens:
        return False
    cursor = content_start
    for index, token in enumerate(tokens):
        gap = text[cursor:token.start]
        if index == 0:
            if gap.strip():
                return False
        elif gap and not _SEPARATOR_RE.fullmatch(gap):
            return False
        cursor = token.end
    return not text[cursor:content_end].strip()


def _overlaps(span: tuple[int, int], spans: Iterable[tuple[int, int]]) -> bool:
    start, end = span
    return any(
        start < other_end and end > other_start for other_start, other_end in spans
    )


def parse_doi_groups(
    text: str,
    *,
    formats: Iterable[str] | None = None,
    excluded_ranges: Sequence[tuple[int, int]] = (),
    story_type: int = 1,
    story_index: int = 0,
) -> list[DoiGroup]:
    """Detect replaceable DOI placeholders in a Word story.

    With no ``formats`` filter, all supported forms are returned. Complete
    bracketed or BibTeX-style groups become one merged citation; otherwise,
    adjacent bare tokens separated by comma, semicolon, or hyphen are merged.
    Ranges occupied by existing Zotero fields can be excluded by the caller.
    """

    selected = {value.casefold() for value in formats or ()}
    invalid = selected - set(DOI_FORMATS)
    if invalid:
        raise ValueError("Unsupported DOI format(s): " + ", ".join(sorted(invalid)))

    tokens = [
        token
        for token in find_doi_tokens(text)
        if not _overlaps((token.start, token.end), excluded_ranges)
    ]
    groups: list[DoiGroup] = []
    covered: list[tuple[int, int]] = []

    wrapper_candidates: list[tuple[int, int, int, int, str]] = []
    for open_start, close_end, open_char in _bracket_pairs(text):
        content_start, content_end = open_start + 1, close_end - 1
        contained = _tokens_in_span(tokens, content_start, content_end)
        if not _only_dois_and_separators(
            text, content_start, content_end, contained
        ):
            continue

        start = open_start
        prefix_match = _BIBTEX_PREFIX_RE.search(text[:open_start])
        placeholder_format = _FORMAT_FOR_OPEN[open_char]
        if prefix_match:
            start = prefix_match.start()
            placeholder_format = "bibtex"
        wrapper_candidates.append(
            (start, close_end, content_start, content_end, placeholder_format)
        )

    # Prefer the widest candidate at a given location (e.g. include ``citep``
    # instead of treating only its braces as the placeholder).
    wrapper_candidates.sort(key=lambda value: (value[0], -(value[1] - value[0])))
    for start, end, content_start, content_end, placeholder_format in wrapper_candidates:
        if _overlaps((start, end), covered) or _overlaps(
            (start, end), excluded_ranges
        ):
            continue
        contained = _tokens_in_span(tokens, content_start, content_end)
        descriptors = {placeholder_format, *(token.format for token in contained)}
        if selected and not selected.intersection(descriptors):
            continue
        groups.append(
            DoiGroup(
                start=start,
                end=end,
                source=text[start:end],
                dois=tuple(token.doi for token in contained),
                format=placeholder_format,
                story_type=story_type,
                story_index=story_index,
            )
        )
        covered.append((start, end))

    remaining = [
        token for token in tokens if not _overlaps((token.start, token.end), covered)
    ]
    index = 0
    while index < len(remaining):
        sequence = [remaining[index]]
        next_index = index + 1
        while next_index < len(remaining):
            previous = sequence[-1]
            candidate = remaining[next_index]
            gap = text[previous.end:candidate.start]
            if not gap or _SEPARATOR_RE.fullmatch(gap) is None:
                break
            sequence.append(candidate)
            next_index += 1

        descriptors = {token.format for token in sequence}
        if not selected or selected.intersection(descriptors):
            start, end = sequence[0].start, sequence[-1].end
            if not _overlaps((start, end), excluded_ranges):
                groups.append(
                    DoiGroup(
                        start=start,
                        end=end,
                        source=text[start:end],
                        dois=tuple(token.doi for token in sequence),
                        format=(
                            sequence[0].format
                            if len(descriptors) == 1
                            else "mixed"
                        ),
                        story_type=story_type,
                        story_index=story_index,
                    )
                )
        index = next_index

    return sorted(groups, key=lambda group: (group.start, group.end))


def parse_parenthesized_dois(text: str) -> list[DoiGroup]:
    """Backward-compatible parser for the original round-bracket behavior."""

    return [
        group
        for group in parse_doi_groups(text, formats={"round"})
        if group.format == "round"
    ]
