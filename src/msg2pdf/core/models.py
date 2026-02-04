"""Data models for msg2pdf."""

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


@dataclass
class Attachment:
    """Represents an email attachment."""

    filename: str
    content_type: str
    data: bytes
    content_id: str | None = None  # For inline images (cid:)
    is_inline: bool = False
    is_msg: bool = False  # Embedded MSG file

    @property
    def size_kb(self) -> float:
        """Return size in KB."""
        return len(self.data) / 1024

    @property
    def size_display(self) -> str:
        """Return human-readable size."""
        size_kb = self.size_kb
        if size_kb < 1024:
            return f"{size_kb:.1f} KB"
        return f"{size_kb / 1024:.1f} MB"


@dataclass
class EmailData:
    """Represents parsed email data."""

    subject: str
    sender: str
    sender_email: str
    to: list[str]
    cc: list[str]
    date: datetime | None
    body_html: str
    body_text: str
    attachments: list[Attachment] = field(default_factory=list)
    source_file: Path | None = None

    @property
    def has_attachments(self) -> bool:
        """Check if email has non-inline attachments."""
        return any(not att.is_inline for att in self.attachments)

    @property
    def inline_attachments(self) -> list[Attachment]:
        """Get inline attachments (images embedded in body)."""
        return [att for att in self.attachments if att.is_inline]

    @property
    def file_attachments(self) -> list[Attachment]:
        """Get regular file attachments."""
        return [att for att in self.attachments if not att.is_inline]

    @property
    def date_display(self) -> str:
        """Return formatted date string."""
        if self.date:
            return self.date.strftime("%B %d, %Y at %I:%M %p")
        return "Unknown"

    @property
    def recipients_display(self) -> str:
        """Return formatted recipients string."""
        return ", ".join(self.to) if self.to else "Unknown"

    @property
    def cc_display(self) -> str:
        """Return formatted CC string."""
        return ", ".join(self.cc) if self.cc else ""
