"""Create the mixed-placeholder DOCX used for the package smoke test."""

from __future__ import annotations

import gc
from pathlib import Path

import win32com.client


def main() -> None:
    output = (
        Path(__file__).resolve().parents[1] / "samples" / "mixed_doi_formats.docx"
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Add()
        document.Content.Text = (
            "Mixed DOI placeholder smoke test\r"
            "Round form: (10.1016/j.compbiomed.2022.106443).\r"
            "Square merged form: [https://doi.org/10.1007/s00701-024-06007-z, "
            "doi: 10.1097/SCS.0000000000003217].\r"
            "Curly form: {10.1055/s-0043-1772170}.\r"
            "Bare hyphen-merged form: 10.1016/j.compbiomed.2022.106443 - "
            "10.1007/s00701-024-06007-z.\r"
            r"BibTeX-like form: \citep{10.20944/preprints202311.0688.v1}."
        )
        title = document.Paragraphs(1).Range
        title.Font.Bold = True
        title.Font.Size = 16
        document.SaveAs2(str(output), FileFormat=16)
        document.Close(SaveChanges=0)
        document = None
    finally:
        if document is not None:
            document.Close(SaveChanges=0)
        try:
            word.Quit()
        finally:
            document = None
            word = None
            gc.collect()
    print(output)


if __name__ == "__main__":
    main()
