"""Create active Zotero citation fields in Microsoft Word documents."""

from .api import ConversionResult, DoiOutcome, cite_document

__all__ = ["ConversionResult", "DoiOutcome", "cite_document"]
__version__ = "0.4.0"
