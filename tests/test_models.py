"""Tests for data models."""

from datetime import datetime

from msg2pdf.core.models import Attachment, EmailData


def test_attachment_size_display_kb():
    """Test attachment size display for KB."""
    att = Attachment(
        filename="test.txt",
        content_type="text/plain",
        data=b"x" * 1024,  # 1 KB
    )
    assert att.size_display == "1.0 KB"


def test_attachment_size_display_mb():
    """Test attachment size display for MB."""
    att = Attachment(
        filename="test.txt",
        content_type="text/plain",
        data=b"x" * (1024 * 1024),  # 1 MB
    )
    assert att.size_display == "1.0 MB"


def test_email_has_attachments():
    """Test has_attachments property."""
    email = EmailData(
        subject="Test",
        sender="test@example.com",
        sender_email="test@example.com",
        to=["recipient@example.com"],
        cc=[],
        date=datetime.now(),
        body_html="<p>Hello</p>",
        body_text="Hello",
        attachments=[
            Attachment(
                filename="file.pdf",
                content_type="application/pdf",
                data=b"pdf data",
                is_inline=False,
            ),
        ],
    )
    assert email.has_attachments is True


def test_email_inline_vs_file_attachments():
    """Test separation of inline and file attachments."""
    email = EmailData(
        subject="Test",
        sender="test@example.com",
        sender_email="test@example.com",
        to=["recipient@example.com"],
        cc=[],
        date=datetime.now(),
        body_html="<p>Hello</p>",
        body_text="Hello",
        attachments=[
            Attachment(
                filename="image.png",
                content_type="image/png",
                data=b"png data",
                content_id="img001",
                is_inline=True,
            ),
            Attachment(
                filename="doc.pdf",
                content_type="application/pdf",
                data=b"pdf data",
                is_inline=False,
            ),
        ],
    )
    assert len(email.inline_attachments) == 1
    assert len(email.file_attachments) == 1
    assert email.inline_attachments[0].filename == "image.png"
    assert email.file_attachments[0].filename == "doc.pdf"


def test_email_date_display():
    """Test date display formatting."""
    email = EmailData(
        subject="Test",
        sender="test@example.com",
        sender_email="test@example.com",
        to=["recipient@example.com"],
        cc=[],
        date=datetime(2024, 1, 15, 14, 30),
        body_html="",
        body_text="",
    )
    assert "January 15, 2024" in email.date_display
    assert "02:30 PM" in email.date_display


def test_email_date_display_none():
    """Test date display when date is None."""
    email = EmailData(
        subject="Test",
        sender="test@example.com",
        sender_email="test@example.com",
        to=[],
        cc=[],
        date=None,
        body_html="",
        body_text="",
    )
    assert email.date_display == "Unknown"
