from __future__ import annotations

from dataclasses import dataclass
import html
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import subprocess
import tempfile
import time
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import ProxyHandler, Request, build_opener
from uuid import uuid4

from .doi import normalize_doi


class ZoteroError(RuntimeError):
    """Raised when Zotero cannot satisfy a citation lookup."""


@dataclass(slots=True)
class ZoteroRecord:
    key: str
    doi: str
    title: str
    uri: str
    item_data: dict[str, Any]
    item_id: int | str | None = None

    @property
    def field_id(self) -> int | str:
        """Best citation ID before Zotero canonicalizes the field on refresh."""

        return self.item_id if self.item_id is not None else self.uri


@dataclass(frozen=True, slots=True)
class DoiResolution:
    doi: str
    selected_key: str
    candidate_keys: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class DoiImportResult:
    doi: str
    status: str
    message: str
    item_key: str | None = None


class ZoteroClient:
    """Small client for Zotero's local HTTP API."""

    def __init__(
        self,
        base_url: str = "http://127.0.0.1:23119",
        library_id: str = "0",
        timeout: float = 15.0,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.library_id = str(library_id)
        self.timeout = timeout
        # Explicitly bypass environment proxies for Zotero's loopback server.
        self._opener = build_opener(ProxyHandler({}))
        self.version = "unknown"

    def _request_json(
        self,
        path: str,
        params: dict[str, Any] | None = None,
        *,
        timeout: float | None = None,
    ) -> tuple[Any, Any]:
        query = f"?{urlencode(params)}" if params else ""
        request = Request(f"{self.base_url}{path}{query}")
        try:
            with self._opener.open(
                request, timeout=self.timeout if timeout is None else timeout
            ) as response:
                self.version = response.headers.get("X-Zotero-Version", self.version)
                return json.load(response), response.headers
        except (OSError, URLError, json.JSONDecodeError) as exc:
            raise ZoteroError(
                f"Zotero's local HTTP API did not respond at {self.base_url}. "
                "Zotero may be open, but the package also requires its local "
                "API/server on port 23119. Check Zotero Settings > Advanced > "
                "Allow other applications on this computer to communicate "
                "with Zotero, then restart Zotero."
            ) from exc

    def wait_until_available(
        self, timeout: float = 12.0, interval: float = 1.0
    ) -> None:
        """Wait briefly for Zotero's desktop API instead of failing instantly."""

        deadline = time.monotonic() + max(0.0, timeout)
        last_error: ZoteroError | None = None
        while True:
            try:
                self._request_json(
                    f"/api/users/{self.library_id}/items",
                    {"format": "json", "limit": 1},
                    timeout=min(2.0, self.timeout),
                )
                return
            except ZoteroError as exc:
                last_error = exc
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            time.sleep(min(max(0.05, interval), remaining))
        raise last_error or ZoteroError(
            f"Zotero's local HTTP API did not respond at {self.base_url}."
        )

    def search_exact_doi(self, doi: str) -> list[dict[str, Any]]:
        normalized = normalize_doi(doi)
        start = 0
        exact: list[dict[str, Any]] = []
        while True:
            payload, headers = self._request_json(
                f"/api/users/{self.library_id}/items",
                {
                    "q": doi,
                    "qmode": "everything",
                    "itemType": "-attachment",
                    "format": "json",
                    "limit": 100,
                    "start": start,
                },
            )
            if not isinstance(payload, list):
                raise ZoteroError("Zotero returned an unexpected item-search response.")
            for item in payload:
                item_doi = item.get("data", {}).get("DOI")
                if isinstance(item_doi, str) and normalize_doi(item_doi) == normalized:
                    exact.append(item)

            total = int(headers.get("Total-Results", len(payload)))
            start += len(payload)
            if not payload or start >= total:
                break

        return sorted(
            exact,
            key=lambda item: (str(item.get("key", "")).casefold(), str(item.get("key", ""))),
        )

    def fetch_csl(self, item_key: str) -> dict[str, Any]:
        payload, _ = self._request_json(
            f"/api/users/{self.library_id}/items/{item_key}",
            {"format": "csljson"},
        )
        if isinstance(payload, list) and len(payload) == 1 and isinstance(payload[0], dict):
            return payload[0]
        if isinstance(payload, dict):
            return payload
        raise ZoteroError(f"Zotero returned invalid CSL JSON for item {item_key}.")

    def resolve_doi_if_present(
        self, doi: str
    ) -> tuple[ZoteroRecord, DoiResolution] | None:
        matches = self.search_exact_doi(doi)
        if not matches:
            return None

        selected = matches[0]
        key = str(selected["key"])
        csl = self.fetch_csl(key)
        uri = csl.get("id")
        if not isinstance(uri, str) or not uri.startswith("http"):
            raise ZoteroError(f"Zotero item {key} did not provide a canonical item URI.")

        record = ZoteroRecord(
            key=key,
            doi=normalize_doi(doi),
            title=str(selected.get("data", {}).get("title", "")),
            uri=uri,
            item_data=csl,
        )
        resolution = DoiResolution(
            doi=normalize_doi(doi),
            selected_key=key,
            candidate_keys=tuple(str(item["key"]) for item in matches),
        )
        return record, resolution

    def resolve_doi(self, doi: str) -> tuple[ZoteroRecord, DoiResolution]:
        result = self.resolve_doi_if_present(doi)
        if result is None:
            raise ZoteroError(f"No exact Zotero record found for DOI {doi}.")
        return result


_ITEM_KEY_RE = re.compile(r"Item key:\s*`?([A-Z0-9]{8})`?", re.IGNORECASE)


def add_doi_with_cli(
    doi: str,
    executable: str = "zotero-cli",
    timeout: float = 180.0,
) -> DoiImportResult:
    """Ask the installed Zotero CLI to fetch and add one DOI idempotently."""

    command = shutil.which(executable)
    if command is None:
        candidate = Path(executable).expanduser()
        if candidate.is_file():
            command = str(candidate.resolve())
    if command is None:
        return DoiImportResult(
            doi=normalize_doi(doi),
            status="failed",
            message=f"Zotero CLI executable was not found: {executable}",
        )

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        completed = subprocess.run(
            [
                command,
                "add",
                "doi",
                normalize_doi(doi),
                "--if-exists",
                "skip",
                "--attach-mode",
                "linked_url",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
            creationflags=creationflags,
        )
    except subprocess.TimeoutExpired:
        return DoiImportResult(
            doi=normalize_doi(doi),
            status="failed",
            message=f"Zotero CLI timed out after {timeout:g} seconds.",
        )
    except OSError as exc:
        return DoiImportResult(
            doi=normalize_doi(doi),
            status="failed",
            message=f"Could not run Zotero CLI: {exc}",
        )

    output = "\n".join(
        value.strip() for value in (completed.stdout, completed.stderr) if value.strip()
    ).strip()
    compact_message = " ".join(output.split()) or "Zotero CLI returned no message."
    key_match = _ITEM_KEY_RE.search(output)
    item_key = key_match.group(1).upper() if key_match else None
    output_lower = output.casefold()

    if "successfully added:" in output_lower:
        status = "added"
    elif "already exists" in output_lower or "existing item" in output_lower:
        status = "already-present"
    elif "doi not found" in output_lower or "no metadata" in output_lower:
        status = "metadata-not-found"
    elif completed.returncode != 0 or any(
        marker in output_lower
        for marker in ("error:", "failed to create item", "input error:")
    ):
        status = "failed"
    else:
        status = "failed"

    return DoiImportResult(
        doi=normalize_doi(doi),
        status=status,
        message=compact_message,
        item_key=item_key,
    )


_CROSSREF_ITEM_TYPES = {
    "journal-article": "journalArticle",
    "proceedings-article": "conferencePaper",
    "book": "book",
    "book-chapter": "bookSection",
    "book-section": "bookSection",
    "dissertation": "thesis",
    "dataset": "dataset",
    "report": "report",
    "report-series": "report",
    "posted-content": "preprint",
    "standard": "standard",
}


def _first(value: Any, default: str = "") -> str:
    if isinstance(value, list) and value:
        return str(value[0])
    if isinstance(value, str):
        return value
    return default


def _crossref_date(metadata: dict[str, Any]) -> str:
    for name in ("published", "published-print", "published-online", "created"):
        parts = metadata.get(name, {}).get("date-parts", [[]])
        if parts and parts[0]:
            return "-".join(str(part) for part in parts[0])
    return ""


def _clean_crossref_abstract(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    without_tags = re.sub(r"<[^>]+>", " ", value)
    return " ".join(html.unescape(without_tags).split())


def _crossref_to_connector_item(doi: str, metadata: dict[str, Any]) -> dict[str, Any]:
    creators = []
    for creator_type, field_name in (("author", "author"), ("editor", "editor")):
        for creator in metadata.get(field_name, []) or []:
            if creator.get("family"):
                creators.append(
                    {
                        "creatorType": creator_type,
                        "firstName": str(creator.get("given", "")),
                        "lastName": str(creator["family"]),
                    }
                )
            elif creator.get("name"):
                creators.append(
                    {"creatorType": creator_type, "name": str(creator["name"])}
                )

    item_type = _CROSSREF_ITEM_TYPES.get(str(metadata.get("type", "")), "document")
    item: dict[str, Any] = {
        "id": uuid4().hex[:8],
        "itemType": item_type,
        "title": _first(metadata.get("title"), normalize_doi(doi)),
        "creators": creators,
        "date": _crossref_date(metadata),
        "DOI": normalize_doi(doi),
        "url": str(metadata.get("URL") or f"https://doi.org/{normalize_doi(doi)}"),
        "publisher": str(metadata.get("publisher", "")),
        "volume": str(metadata.get("volume", "")),
        "issue": str(metadata.get("issue", "")),
        "pages": str(metadata.get("page", "")),
        "ISSN": _first(metadata.get("ISSN")),
        "publicationTitle": _first(metadata.get("container-title")),
        "abstractNote": _clean_crossref_abstract(metadata.get("abstract")),
        "language": str(metadata.get("language", "")),
        "libraryCatalog": "Crossref",
        "tags": [],
        "notes": [],
        "attachments": [],
    }
    return {key: value for key, value in item.items() if value not in ("", None)}


def add_doi_to_local_zotero(
    doi: str,
    base_url: str = "http://127.0.0.1:23119",
    timeout: float = 30.0,
) -> DoiImportResult:
    """Fetch Crossref metadata and save it through Zotero's Connector server.

    This is an API-key-free fallback for installations where ``zotero-cli`` is
    configured in local-only mode. It writes directly to the running desktop
    Zotero library and deliberately creates no attachment.
    """

    normalized = normalize_doi(doi)
    crossref_url = f"https://api.crossref.org/works/{quote(normalized, safe='')}"
    crossref_request = Request(
        crossref_url,
        headers={
            "Accept": "application/json",
            "User-Agent": "zotero-word-citations/0.2 (DOI metadata import)",
        },
    )
    try:
        with build_opener().open(crossref_request, timeout=timeout) as response:
            metadata = json.load(response).get("message", {})
    except HTTPError as exc:
        if exc.code == 404:
            return DoiImportResult(
                doi=normalized,
                status="metadata-not-found",
                message=f"DOI not found on Crossref: {normalized}",
            )
        return DoiImportResult(
            doi=normalized,
            status="failed",
            message=f"Crossref returned HTTP {exc.code} for {normalized}.",
        )
    except (OSError, URLError, json.JSONDecodeError) as exc:
        return DoiImportResult(
            doi=normalized,
            status="failed",
            message=f"Could not fetch Crossref metadata: {exc}",
        )

    if not isinstance(metadata, dict) or not metadata.get("title"):
        return DoiImportResult(
            doi=normalized,
            status="metadata-not-found",
            message=f"Crossref returned no usable metadata for {normalized}.",
        )

    payload = {
        "sessionID": f"zotero-word-citations-{uuid4().hex}",
        "uri": f"https://doi.org/{normalized}",
        "items": [_crossref_to_connector_item(normalized, metadata)],
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    connector_request = Request(
        f"{base_url.rstrip('/')}/connector/saveItems",
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "X-Zotero-Connector-API-Version": "3",
        },
    )
    try:
        opener = build_opener(ProxyHandler({}))
        with opener.open(connector_request, timeout=timeout) as response:
            if response.status not in (200, 201):
                raise ZoteroError(
                    f"Zotero Connector returned HTTP {response.status}."
                )
    except (OSError, URLError, ZoteroError) as exc:
        return DoiImportResult(
            doi=normalized,
            status="failed",
            message=f"Could not save metadata to local Zotero: {exc}",
        )

    return DoiImportResult(
        doi=normalized,
        status="added",
        message=(
            "Metadata fetched from Crossref and added directly to the running "
            "Zotero desktop library (no attachment)."
        ),
    )


def discover_zotero_data_dir(explicit: Path | None = None) -> Path | None:
    """Locate the common Zotero data directory used for numeric item IDs."""

    candidates: list[Path] = []
    if explicit is not None:
        candidates.append(explicit.expanduser())
    env_path = os.environ.get("ZOTERO_DATA_DIR")
    if env_path:
        candidates.append(Path(env_path).expanduser())
    candidates.append(Path.home() / "Zotero")

    for candidate in candidates:
        if (candidate / "zotero.sqlite").is_file():
            return candidate.resolve()
    return None


def resolve_numeric_item_ids(data_dir: Path | None, keys: Iterable[str]) -> dict[str, int]:
    """Read Zotero internal item IDs from a disposable SQLite snapshot.

    Zotero keeps its live database locked. Copying the database plus WAL files
    gives us a read-only, short-lived snapshot without touching Zotero's data.
    Missing IDs are safe: fields retain canonical URIs and Zotero resolves them
    during Refresh.
    """

    wanted = sorted(set(keys))
    if data_dir is None or not wanted:
        return {}
    source_db = data_dir / "zotero.sqlite"
    if not source_db.is_file():
        return {}

    try:
        with tempfile.TemporaryDirectory(prefix="zotero-word-cite-") as temp_name:
            temp_dir = Path(temp_name)
            for suffix in ("", "-wal", "-shm"):
                source = Path(f"{source_db}{suffix}")
                if source.exists():
                    shutil.copy2(source, temp_dir / source.name)

            snapshot = temp_dir / "zotero.sqlite"
            connection = sqlite3.connect(snapshot)
            try:
                placeholders = ",".join("?" for _ in wanted)
                rows = connection.execute(
                    f"SELECT itemID, key FROM items WHERE key IN ({placeholders})",
                    wanted,
                ).fetchall()
            finally:
                # sqlite3.Connection's context manager does not close the
                # handle, which prevents TemporaryDirectory cleanup on Windows.
                connection.close()
            return {str(key): int(item_id) for item_id, key in rows}
    except (OSError, sqlite3.Error):
        return {}
