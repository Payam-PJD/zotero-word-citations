# Contributing

This project targets Windows because it uses Microsoft Word COM automation.

1. Create a virtual environment with Python 3.10 or newer.
2. Install the project with `python -m pip install -e .`.
3. Run `python -m unittest discover -s tests -v`.
4. Keep test documents and Zotero libraries out of commits; use synthetic DOI
   values in unit tests whenever possible.

Changes that touch Word field insertion should be checked in a disposable copy
of a `.docx`. Confirm that existing Zotero citation and bibliography fields are
unchanged, save and reopen the output, and run **Zotero > Refresh** before
considering the change verified.
