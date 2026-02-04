"""Custom exceptions for msg2pdf."""


class MSG2PDFError(Exception):
    """Base exception for msg2pdf errors."""

    pass


class MSGParseError(MSG2PDFError):
    """Error parsing MSG file."""

    pass


class PDFRenderError(MSG2PDFError):
    """Error rendering PDF."""

    pass


class AttachmentError(MSG2PDFError):
    """Error handling attachment."""

    pass


class BatchProcessingError(MSG2PDFError):
    """Error during batch processing."""

    def __init__(self, message: str, failed_files: list[tuple[str, str]] | None = None):
        super().__init__(message)
        self.failed_files = failed_files or []
