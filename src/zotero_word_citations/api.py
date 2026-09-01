from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
import os
from pathlib import Path
import sys
import time
from typing import Callable, Iterable, Mapping, Sequence
from uuid import uuid4

from .doi import DoiGroup, normalize_doi
from .word import (
    DEFAULT_STYLE_ID,
    InsertedCitation,
    WORD_STORY_NAMES,
    WordError,
    convert_word_document,
    scan_word_doi_like_values,
    scan_word_document,
)
from .zotero import (
    DoiResolution,
    ZoteroClient,
    ZoteroError,
    ZoteroRecord,
    add_doi_to_local_zotero,
    add_doi_with_cli,
    discover_zotero_data_dir,
    resolve_numeric_item_ids,
)


ProgressCallback = Callable[[str], None]


@dataclass(frozen=True, slots=True)
class DoiOutcome:
    doi: str
    status: str
    title: str = ""
    selected_key: str = ""
    candidate_keys: tuple[str, ...] = ()
    message: str = ""


@dataclass(frozen=True, slots=True)
class PlaceholderOutcome:
    group: DoiGroup
    status: str
    message: str


@dataclass(frozen=True, slots=True)
class ConversionResult:
    input_path: Path
    output_path: Path | None
    report_path: Path
    discovered_groups: tuple[DoiGroup, ...]
    inserted_citations: tuple[InsertedCitation, ...]
    doi_outcomes: tuple[DoiOutcome, ...]
    placeholder_outcomes: tuple[PlaceholderOutcome, ...]
    created_preferences: bool
    warnings: tuple[str, ...] = ()
    left_in_prose: tuple[str, ...] = ()

    @property
    def discovered_dois(self) -> tuple[str, ...]:
        return tuple(
            dict.fromkeys(
                doi for group in self.discovered_groups for doi in group.dois
            )
        )


def _default_output(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_zotero_cited.docx")


def _default_report(input_path: Path) -> Path:
    return input_path.with_name("dois.txt")


def _progress_handler(
    *, verbose: bool, progress: ProgressCallback | None
) -> ProgressCallback:
    if progress is not None:
        return progress
    if verbose:
        return _safe_console_print
    return lambda _message: None


def _safe_console_print(message: str) -> None:
    """Print progress without letting a legacy Windows code page abort work."""

    value = str(message)
    try:
        print(value)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
        safe = value.encode(encoding, errors="replace").decode(
            encoding, errors="replace"
        )
        print(safe)


def _single_line(value: str) -> str:
    return " ".join(value.replace("\x07", " ").split())


def write_doi_report(
    path: Path,
    *,
    input_path: Path,
    output_path: Path | None,
    groups: Sequence[DoiGroup],
    outcomes: Mapping[str, DoiOutcome],
    placeholder_outcomes: Sequence[PlaceholderOutcome] = (),
    formats: Iterable[str] | None = None,
    warnings: Sequence[str] = (),
    left_in_prose: Sequence[str] = (),
) -> None:
    """Atomically write the human-readable DOI audit report."""

    unique_dois = list(dict.fromkeys(doi for group in groups for doi in group.dois))
    status_counts: dict[str, int] = {}
    for doi in unique_dois:
        status = outcomes.get(doi, DoiOutcome(doi, "discovered")).status
        status_counts[status] = status_counts.get(status, 0) + 1

    lines = [
        "Zotero Word DOI conversion report",
        "=" * 34,
        f"Generated: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"Input: {input_path}",
        f"Output: {output_path if output_path is not None else '(not created)'}",
        f"Formats: {', '.join(formats) if formats else 'auto (all supported forms)'}",
        f"Placeholders discovered: {len(groups)}",
        f"Unique DOI-like strings: {len(unique_dois)}",
    ]
    if status_counts:
        lines.append(
            "Status summary: "
            + ", ".join(f"{key}={value}" for key, value in sorted(status_counts.items()))
        )

    lines.extend(["", "DOI results", "-----------"])
    if not unique_dois:
        lines.append("(none)")
    for number, doi in enumerate(unique_dois, start=1):
        outcome = outcomes.get(doi, DoiOutcome(doi=doi, status="discovered"))
        lines.append(f"{number}. [{outcome.status.upper()}] {doi}")
        if outcome.title:
            lines.append(f"   Title: {_single_line(outcome.title)}")
        if outcome.selected_key:
            lines.append(f"   Selected Zotero key: {outcome.selected_key}")
        if len(outcome.candidate_keys) > 1:
            lines.append(f"   Candidate keys: {', '.join(outcome.candidate_keys)}")
        if outcome.message:
            lines.append(f"   Note: {_single_line(outcome.message)}")
        occurrences = sum(group.dois.count(doi) for group in groups)
        lines.append(f"   Occurrences: {occurrences}")

    lines.extend(["", "Placeholder actions", "-------------------"])
    if not groups:
        lines.append("(none)")
    outcome_by_group = {item.group: item for item in placeholder_outcomes}
    for number, group in enumerate(groups, start=1):
        action = outcome_by_group.get(
            group,
            PlaceholderOutcome(group, "discovered", "Awaiting Zotero resolution."),
        )
        story_name = WORD_STORY_NAMES.get(group.story_type, f"story {group.story_type}")
        lines.append(
            f"{number}. [{action.status.upper()}] {story_name} #{group.story_index + 1}, "
            f"characters {group.start}-{group.end}, format={group.format}"
        )
        lines.append(f"   Source: {_single_line(group.source)}")
        lines.append(f"   DOIs: {', '.join(group.dois)}")
        lines.append(f"   Note: {_single_line(action.message)}")

    lines.extend(["", "Warnings", "--------"])
    if not warnings:
        lines.append("(none)")
    else:
        for number, warning in enumerate(warnings, start=1):
            lines.append(f"{number}. {_single_line(warning)}")

    lines.extend(["", "LEFT IN PROSE", "-------------"])
    if not left_in_prose:
        lines.append("(none)")
    else:
        lines.append(
            "These DOI-like values remain outside protected Word fields and "
            "were not deliberately skipped for unresolved metadata:"
        )
        for doi in left_in_prose:
            lines.append(f"- {doi}")

    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    temporary.write_text("\n".join(lines) + "\n", encoding="utf-8")
    os.replace(temporary, path)


def _outcome_from_record(
    doi: str,
    record: ZoteroRecord,
    resolution: DoiResolution,
    status: str,
    message: str = "",
) -> DoiOutcome:
    return DoiOutcome(
        doi=doi,
        status=status,
        title=record.title,
        selected_key=record.key,
        candidate_keys=resolution.candidate_keys,
        message=message,
    )


def cite_document(
    input_path: str | Path,
    output_path: str | Path | None = None,
    *,
    report_path: str | Path | None = None,
    formats: Iterable[str] | None = None,
    add_missing: bool = True,
    force: bool = False,
    zotero_url: str = "http://127.0.0.1:23119",
    library_id: str = "0",
    zotero_data_dir: str | Path | None = None,
    style_id: str = DEFAULT_STYLE_ID,
    zotero_cli: str | None = None,
    import_timeout: float = 180.0,
    sync_timeout: float = 60.0,
    verbose: bool = True,
    progress: ProgressCallback | None = None,
) -> ConversionResult:
    """Discover DOI placeholders, resolve/import them, and create Zotero fields.

    This is the notebook-friendly API behind the command-line interface. The
    source document is never overwritten. Groups containing any unresolved DOI
    remain unchanged, while other groups are converted.
    """

    emit = _progress_handler(verbose=verbose, progress=progress)
    source = Path(input_path).expanduser().resolve()
    destination = (
        Path(output_path).expanduser().absolute()
        if output_path is not None
        else _default_output(source)
    )
    report = (
        Path(report_path).expanduser().absolute()
        if report_path is not None
        else _default_report(source)
    )
    selected_formats = tuple(dict.fromkeys(value.casefold() for value in formats or ()))

    if not source.is_file():
        raise WordError(f"Input document does not exist: {source}")
    if source.suffix.casefold() != ".docx" or destination.suffix.casefold() != ".docx":
        raise WordError("Both input and output must use the .docx extension.")
    if source == destination.resolve():
        raise WordError("The output must differ from the input; source files are preserved.")
    if destination.exists() and not force:
        raise WordError(f"Output already exists: {destination}. Use --force to replace it.")
    if report.resolve() in {source, destination.resolve()}:
        raise WordError("The DOI report path must differ from the input and output paths.")

    run_warnings: list[str] = []
    emit(f"[1/4] Scanning Word document: {source.name}")
    groups = scan_word_document(
        source,
        formats=selected_formats or None,
        warnings=run_warnings,
    )
    unique_dois = list(dict.fromkeys(doi for group in groups for doi in group.dois))
    emit(
        f"      Found {len(groups)} placeholder(s) containing "
        f"{len(unique_dois)} unique DOI-like string(s)."
    )

    outcomes: dict[str, DoiOutcome] = {
        doi: DoiOutcome(doi=doi, status="discovered") for doi in unique_dois
    }
    write_doi_report(
        report,
        input_path=source,
        output_path=None,
        groups=groups,
        outcomes=outcomes,
        formats=selected_formats,
        warnings=run_warnings,
    )
    emit(f"      Discovery report initialized: {report}")

    if not groups:
        emit("[DONE] No DOI placeholders were found; no Word output was created.")
        return ConversionResult(
            input_path=source,
            output_path=None,
            report_path=report,
            discovered_groups=(),
            inserted_citations=(),
            doi_outcomes=(),
            placeholder_outcomes=(),
            created_preferences=False,
            warnings=tuple(run_warnings),
        )

    emit("[2/4] Resolving DOI metadata in Zotero.")
    client = ZoteroClient(zotero_url, library_id)
    emit("      Waiting for Zotero's local API (up to 12 seconds).")
    client.wait_until_available(timeout=12.0)
    emit(f"      Connected to Zotero {client.version}.")
    records: dict[str, ZoteroRecord] = {}
    for index, doi in enumerate(unique_dois, start=1):
        emit(f"      ({index}/{len(unique_dois)}) {doi}")
        try:
            present = client.resolve_doi_if_present(doi)
        except ZoteroError as exc:
            outcomes[doi] = DoiOutcome(
                doi=doi,
                status="failed",
                message=str(exc),
            )
            emit(f"        [FAILED] {exc}")
            continue
        if present is not None:
            record, resolution = present
            records[doi] = record
            outcomes[doi] = _outcome_from_record(
                doi, record, resolution, "already-present"
            )
            emit(f"        [FOUND] Zotero key {record.key}: {record.title}")
            continue

        if not add_missing:
            outcomes[doi] = DoiOutcome(
                doi=doi,
                status="failed",
                message="Not present in Zotero; automatic import was disabled.",
            )
            emit("        [SKIP] Missing and --no-add-missing was selected.")
            continue

        emit(
            "        [ADD] Not in the local library; fetching metadata for "
            "the running Zotero desktop app."
        )
        local_import = add_doi_to_local_zotero(
            doi,
            base_url=zotero_url,
            timeout=min(import_timeout, 60.0),
        )
        imported = local_import
        if (
            local_import.status not in {"added", "already-present"}
            and zotero_cli
        ):
            emit(
                "        [FALLBACK] Local metadata import did not succeed; "
                "trying the optional zotero-cli cloud writer."
            )
            cli_import = add_doi_with_cli(
                doi, executable=zotero_cli, timeout=import_timeout
            )
            if cli_import.status in {"added", "already-present"}:
                imported = type(cli_import)(
                    doi=cli_import.doi,
                    status=cli_import.status,
                    message=(
                        f"Local import did not succeed ({local_import.message}) "
                        f"zotero-cli succeeded: {cli_import.message}"
                    ),
                    item_key=cli_import.item_key,
                )
            elif local_import.status == "metadata-not-found":
                imported = type(local_import)(
                    doi=local_import.doi,
                    status=local_import.status,
                    message=(
                        f"{local_import.message} Optional zotero-cli also did "
                        f"not add it: {cli_import.message}"
                    ),
                    item_key=local_import.item_key,
                )
            else:
                imported = type(local_import)(
                    doi=local_import.doi,
                    status="failed",
                    message=(
                        f"Local import failed: {local_import.message} "
                        f"Optional zotero-cli also failed: {cli_import.message}"
                    ),
                    item_key=cli_import.item_key or local_import.item_key,
                )
        if imported.status == "metadata-not-found":
            outcomes[doi] = DoiOutcome(
                doi=doi,
                status="metadata-not-found",
                message=imported.message,
            )
            emit("        [NOT FOUND] Crossref/Zotero returned no metadata.")
            continue
        if imported.status == "failed":
            outcomes[doi] = DoiOutcome(
                doi=doi, status="failed", message=imported.message
            )
            emit(f"        [FAILED] {imported.message}")
            continue

        emit("        [SYNC] Metadata added; waiting for the desktop library.")
        deadline = time.monotonic() + max(0.0, sync_timeout)
        last_notice = 0.0
        try:
            synced = client.resolve_doi_if_present(doi)
        except ZoteroError:
            synced = None
        while synced is None and time.monotonic() < deadline:
            remaining = max(0.0, deadline - time.monotonic())
            if time.monotonic() - last_notice >= 10:
                emit(f"        [SYNC] Up to {remaining:.0f}s remaining...")
                last_notice = time.monotonic()
            time.sleep(min(2.0, remaining))
            try:
                synced = client.resolve_doi_if_present(doi)
            except ZoteroError:
                synced = None

        if synced is None:
            outcomes[doi] = DoiOutcome(
                doi=doi,
                status="failed",
                selected_key=imported.item_key or "",
                message=(
                    "Metadata was added but did not become visible "
                    f"in the local Zotero library within {sync_timeout:g} seconds. "
                    "Run Zotero Sync and try again."
                ),
            )
            emit("        [FAILED] Added, but local Zotero did not expose it in time.")
            continue

        record, resolution = synced
        records[doi] = record
        outcomes[doi] = _outcome_from_record(
            doi,
            record,
            resolution,
            "metadata-fetched-and-added",
            imported.message,
        )
        emit(f"        [ADDED] Zotero key {record.key}: {record.title}")

        write_doi_report(
            report,
            input_path=source,
            output_path=None,
            groups=groups,
            outcomes=outcomes,
            formats=selected_formats,
            warnings=run_warnings,
        )

    data_dir = discover_zotero_data_dir(
        Path(zotero_data_dir) if zotero_data_dir is not None else None
    )
    numeric_ids = resolve_numeric_item_ids(data_dir, (record.key for record in records.values()))
    for record in records.values():
        record.item_id = numeric_ids.get(record.key)

    convertible: list[DoiGroup] = []
    placeholder_outcomes: list[PlaceholderOutcome] = []
    for group in groups:
        unresolved = [doi for doi in group.dois if doi not in records]
        if unresolved:
            placeholder_outcomes.append(
                PlaceholderOutcome(
                    group=group,
                    status="skipped",
                    message=(
                        "Left unchanged because these DOI values were unresolved: "
                        + ", ".join(unresolved)
                    ),
                )
            )
        else:
            convertible.append(group)
            placeholder_outcomes.append(
                PlaceholderOutcome(
                    group=group,
                    status="cited",
                    message="Replaced by one active Zotero citation field.",
                )
            )

    inserted: list[InsertedCitation] = []
    created_preferences = False
    created_output: Path | None = None
    if convertible:
        emit(
            f"[3/4] Replacing {len(convertible)} placeholder(s) with active Zotero fields."
        )
        inserted, created_preferences = convert_word_document(
            input_path=source,
            output_path=destination,
            groups=convertible,
            records_by_doi=records,
            zotero_version=client.version,
            style_id=style_id,
            force=force,
        )
        created_output = destination
        emit(f"      Verified {len(inserted)} persisted Word ADDIN field(s).")
    else:
        emit("[3/4] No placeholder had complete metadata; Word output was not created.")

    left_in_prose: tuple[str, ...] = ()
    if created_output is not None:
        remaining = scan_word_doi_like_values(created_output)
        deliberately_skipped = Counter(
            doi
            for item in placeholder_outcomes
            if item.status == "skipped"
            for doi in item.group.dois
        )
        unexpected: list[str] = []
        for doi in remaining:
            if deliberately_skipped[doi] > 0:
                deliberately_skipped[doi] -= 1
            else:
                unexpected.append(doi)
        left_in_prose = tuple(dict.fromkeys(unexpected))
        if left_in_prose:
            warning = (
                f"{len(left_in_prose)} DOI-like value(s) remain outside active "
                "fields and were not deliberately skipped. Review LEFT IN PROSE."
            )
            run_warnings.append(warning)
            emit(f"      [WARNING] {warning}")

    emit("[4/4] Finalizing dois.txt audit report.")
    write_doi_report(
        report,
        input_path=source,
        output_path=created_output,
        groups=groups,
        outcomes=outcomes,
        placeholder_outcomes=placeholder_outcomes,
        formats=selected_formats,
        warnings=run_warnings,
        left_in_prose=left_in_prose,
    )
    if created_output is not None:
        emit(f"[DONE] Created: {created_output}")
        emit("       Open it in Word and use Zotero > Refresh to finalize formatting.")
    emit(f"       Report: {report}")

    return ConversionResult(
        input_path=source,
        output_path=created_output,
        report_path=report,
        discovered_groups=tuple(groups),
        inserted_citations=tuple(inserted),
        doi_outcomes=tuple(outcomes[doi] for doi in unique_dois),
        placeholder_outcomes=tuple(placeholder_outcomes),
        created_preferences=created_preferences,
        warnings=tuple(run_warnings),
        left_in_prose=left_in_prose,
    )
