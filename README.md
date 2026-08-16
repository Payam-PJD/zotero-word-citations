# Zotero Word Citations

Windows package and CLI that discovers DOI placeholders throughout a `.docx`,
writes a `dois.txt` audit beside the input, resolves each DOI against Zotero,
adds missing records through the installed Zotero CLI, and replaces resolvable
placeholders with genuine Word `ADDIN ZOTERO_ITEM CSL_CITATION` fields.

The source document is never overwritten. Existing active Zotero citations and
bibliographies are excluded in full: neither their hidden field-code JSON nor
their displayed results are scanned or edited.

> **Alpha software:** work on a copy of important documents and review the
> generated file before adopting it. This independent project is not affiliated
> with or endorsed by Zotero or Microsoft.

## Requirements

- Windows with desktop Microsoft Word installed
- Zotero running with its local HTTP API enabled (default port `23119`)
- `zotero-cli` installed if cloud-backed imports are desired; local-only CLI
  installations automatically fall back to Zotero's desktop Connector API
- Python 3.10+

## Installation

Install the latest version directly from GitHub:

```powershell
py -m pip install --upgrade "git+https://github.com/Payam-PJD/zotero-word-citations.git"
```

If `py` is unavailable, use `python` in the same command. To install from a
local clone instead:

```powershell
git clone https://github.com/Payam-PJD/zotero-word-citations.git
cd zotero-word-citations
py -m pip install .
```

Verify the installation:

```powershell
py -m zotero_word_citations --help
zotero-word-cite --help
```

The second form requires Python's Scripts folder to be on `PATH`; the first
form works without that extra entry.

## Use from a terminal

```powershell
zotero-word-cite "C:\research\my-document.docx"
```

The equivalent command that does not depend on the Scripts folder being on
`PATH` is:

```powershell
py -m zotero_word_citations "C:\research\my-document.docx"
```

The command prints four progress stages and creates:

- `my-document_zotero_cited.docx`
- `dois.txt`

Open the output in Word and select **Zotero > Refresh** to let Zotero/citeproc
finalize numbering, sorting, prefixes, and the selected citation style.

Use an explicit output or report path when needed:

```powershell
zotero-word-cite input.docx --output cited.docx --report input-dois.txt
```

An existing output is protected unless `--force` is supplied.

## DOI forms detected by default

With no `--format` option, the detector processes every supported form:

```text
10.1000/example
doi: 10.1000/example
https://doi.org/10.1000/example
(10.1000/example)
[10.1000/one, 10.1000/two]
{10.1000/one; 10.1000/two}
10.1000/one - 10.1000/two
\citep{10.1000/one,10.1000/two}
```

Comma-, semicolon-, and hyphen-separated DOI sequences become one merged
Zotero citation field. Parentheses, brackets, braces, a DOI URL/prefix, and a
recognized LaTeX citation command are removed as part of the placeholder.
Sentence punctuation outside the placeholder is retained.

To restrict processing, repeat `--format` with one or more of:

```text
bare  url  doi-prefix  round  square  curly  bibtex
```

For example:

```powershell
zotero-word-cite input.docx --format round --format bibtex
```

The detector uses the standard `10.<registrant>/<suffix>` DOI shape plus broad
legacy-suffix support. Syntactically DOI-like strings are reported even if no
metadata service recognizes them.

## Missing and duplicate records

- Existing DOI matches are case-insensitive and otherwise exact.
- Duplicate Zotero records are sorted alphabetically by their eight-character
  Zotero item key; the first key is cited.
- A missing DOI is first passed to `zotero-cli add doi` in idempotent `skip`
  mode.
- If the CLI is unavailable for writes (including local-only mode), metadata is
  fetched from Crossref and saved directly through the running Zotero desktop
  Connector API. This fallback creates no attachment and requires no cloud API
  key.
- After either import path, the package waits for the record to become visible
  in the desktop library before creating its field. Adjust this with
  `--sync-timeout SECONDS`.
- Use `--no-add-missing` when the library must remain read-only.
- If one DOI in a merged placeholder cannot be resolved, that entire placeholder
  is left unchanged so no source information is lost. Other complete
  placeholders are still converted.

The CLI path uses `linked_url` attachment mode, avoiding automatic PDF
downloads. The local Connector fallback adds metadata only.

## `dois.txt` audit

The report is initialized immediately after discovery and finalized at the end.
It records every unique DOI and occurrence, the source placeholder and location,
selected Zotero item key/title, duplicate candidates, and one of these outcomes:

- `ALREADY-PRESENT`
- `METADATA-FETCHED-AND-ADDED`
- `METADATA-NOT-FOUND`
- `FAILED`

Every placeholder is separately marked `CITED` or `SKIPPED`.

## Use from Jupyter/IPython

Install into the active notebook kernel if needed:

```python
%pip install --upgrade "git+https://github.com/Payam-PJD/zotero-word-citations.git"
```

Restart the kernel after a first installation, then call the Python API:

```python
from zotero_word_citations import cite_document

result = cite_document(
    r"C:\research\draft.docx",
    output_path=r"C:\research\draft-cited.docx",
    report_path=r"C:\research\dois.txt",
    force=True,
)

print(result.output_path)
print(result.report_path)
for item in result.doi_outcomes:
    print(item.doi, item.status, item.selected_key)
```

Notebook progress messages are enabled by default. Set `verbose=False` for a
quiet run, or pass `progress=my_callback` to route messages to a widget/logger.

## Use from a Word VBA shortcut

The repository includes
[`vba/ZoteroWordCitations.bas`](vba/ZoteroWordCitations.bas). Its
`ZoteroCiteDOIsInActiveDocument` macro saves the active `.docx`, runs the same
Python package in a visible terminal, and opens the generated
`*_zotero_cited.docx`. It never overwrites the source document.

### Add the macro to Word

1. Install the Python package using the instructions above.
2. Download `ZoteroWordCitations.bas` from this repository.
3. In Word, press **Alt+F11** to open the Visual Basic Editor.
4. Select the **Normal** project so the macro is available to every document.
5. Choose **File > Import File**, select `ZoteroWordCitations.bas`, and save the
   Normal template when Word asks.
6. Open **File > Options > Quick Access Toolbar**.
7. Set **Choose commands from** to **Macros**, add
   `Normal.ZoteroWordCitations.ZoteroCiteDOIsInActiveDocument`, optionally
   rename it and choose an icon, then select **OK**.

Open and save a `.docx`, make it the active document, and select the new Quick
Access button. Keep Zotero running. The macro waits for conversion, opens the
new file when successful, and tells you to use **Zotero > Refresh**.

The macro starts Python with the Windows launcher command `py`. If Word reports that Python
cannot be found, open the imported module in the Visual Basic Editor and change
this line to the full path of the Python installation where the package was
installed:

```vb
Private Const PYTHON_COMMAND As String = "C:\Users\your-name\AppData\Local\Programs\Python\Python312\python.exe"
```

The VBA module is plain source and is intentionally unsigned. Import macros
only from a source you trust. Organization-managed Word installations may
require an administrator-approved trusted location or a locally signed macro.

## Safety and field behavior

- The input path and output path must differ.
- Output is written through a temporary DOCX and committed only after Word is
  reopened and every inserted field is verified as type 81 (`ADDIN`).
- Zotero document preferences are preserved. If absent, valid field preferences
  are created using American Medical Association as the provisional style.
- Main text, footnotes, endnotes, text frames, and headers/footers are scanned.
- Every active Zotero citation or bibliography is treated as a protected
  envelope. This includes its hidden instruction/JSON, field markers, and
  displayed result.
- Other Word fields remain protected unless their entire displayed result is
  exactly one DOI placeholder. This narrow exception lets autoformatted DOI
  hyperlinks be replaced by a Zotero field while removing the complete old
  hyperlink field.
- Literal DOI placeholders before or after existing fields use Word's real
  character coordinates, so hidden Zotero JSON cannot shift a replacement onto
  the wrong text.
- A second overlap check runs immediately before insertion and aborts the
  conversion if a requested replacement touches any existing Zotero field.

## Development

Run the test suite on Windows:

```powershell
py -m pip install -e .
py -m unittest discover -s tests -v
```

GitHub Actions repeats the unit tests on Python 3.10 through 3.13. See
[`CONTRIBUTING.md`](CONTRIBUTING.md) for Word-field verification guidance.

## License

MIT. See [`LICENSE`](LICENSE).
