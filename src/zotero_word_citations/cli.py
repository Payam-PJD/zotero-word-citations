from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .api import cite_document
from .doi import DOI_FORMATS
from .word import DEFAULT_STYLE_ID, WordError
from .zotero import ZoteroError


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="zotero-word-cite",
        description=(
            "Find DOI placeholders throughout a DOCX, record every DOI in "
            "dois.txt, add missing Zotero records, and replace resolvable "
            "placeholders with active Zotero citation fields."
        ),
    )
    parser.add_argument("input", type=Path, help="Source .docx file (never overwritten).")
    parser.add_argument("-o", "--output", type=Path, help="Output .docx path.")
    parser.add_argument(
        "--report",
        type=Path,
        help="DOI audit report path (default: dois.txt beside the input).",
    )
    parser.add_argument(
        "--format",
        dest="formats",
        action="append",
        choices=DOI_FORMATS,
        help=(
            "Only process a detected form; repeat for more than one. If omitted, "
            "all bare, URL, bracketed, and BibTeX-like forms are processed."
        ),
    )
    parser.add_argument(
        "--no-add-missing",
        dest="add_missing",
        action="store_false",
        help="Leave DOI placeholders unchanged when the item is absent from Zotero.",
    )
    parser.set_defaults(add_missing=True)
    parser.add_argument(
        "--force", action="store_true", help="Replace an existing output file."
    )
    parser.add_argument(
        "--zotero-url",
        default="http://127.0.0.1:23119",
        help="Zotero local API base URL.",
    )
    parser.add_argument(
        "--library-id", default="0", help="Zotero user-library ID for the local API."
    )
    parser.add_argument(
        "--zotero-data-dir",
        type=Path,
        help="Optional Zotero data directory used to obtain internal numeric item IDs.",
    )
    parser.add_argument(
        "--zotero-cli",
        default=None,
        help=(
            "Optional Zotero CLI executable for a secondary cloud-write "
            "fallback. Omit for normal local desktop use."
        ),
    )
    parser.add_argument(
        "--import-timeout",
        type=float,
        default=180.0,
        help="Seconds allowed for each zotero-cli metadata import (default: 180).",
    )
    parser.add_argument(
        "--sync-timeout",
        type=float,
        default=60.0,
        help="Seconds to wait for a newly added record to appear locally (default: 60).",
    )
    parser.add_argument(
        "--style-id",
        default=DEFAULT_STYLE_ID,
        help="CSL style used only when the document has no Zotero preferences.",
    )
    parser.add_argument(
        "--quiet", action="store_true", help="Suppress progress messages."
    )
    return parser


def run(args: argparse.Namespace) -> int:
    result = cite_document(
        args.input,
        args.output,
        report_path=args.report,
        formats=args.formats,
        add_missing=args.add_missing,
        force=args.force,
        zotero_url=args.zotero_url,
        library_id=args.library_id,
        zotero_data_dir=args.zotero_data_dir,
        style_id=args.style_id,
        zotero_cli=args.zotero_cli,
        import_timeout=args.import_timeout,
        sync_timeout=args.sync_timeout,
        verbose=not args.quiet,
    )
    return 0 if result.output_path is not None or not result.discovered_groups else 2


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return run(args)
    except (ValueError, WordError, ZoteroError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
