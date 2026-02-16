"""MarkItWrite - AI virtual clipboard for pasting images into any document format."""

from markitwrite._base_writer import DocumentWriter, WriteResult
from markitwrite._markitwrite import MarkItWrite

__all__ = ["MarkItWrite", "DocumentWriter", "WriteResult"]
__version__ = "0.2.0"
