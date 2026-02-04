"""Tests for MSG to PDF converter."""

import pytest
from pathlib import Path
from datetime import datetime
from unittest.mock import Mock, patch

from msg2pdf.core.converter import MSGToPDFConverter
from msg2pdf.core.models import EmailData, Attachment


class TestMSGToPDFConverter:
    """Tests for MSGToPDFConverter class."""

    @pytest.fixture
    def converter(self):
        """Create a converter instance."""
        return MSGToPDFConverter()

    @pytest.fixture
    def mock_email(self):
        """Create a mock email data object."""
        return EmailData(
            subject="Test Subject",
            sender="sender@example.com",
            sender_email="sender@example.com",
            to=["recipient@example.com"],
            cc=[],
            date=datetime(2024, 1, 15, 14, 30),
            body_html="<p>Hello World</p>",
            body_text="Hello World",
            attachments=[
                Attachment(
                    filename="document.pdf",
                    content_type="application/pdf",
                    data=b"PDF_DATA",
                    is_inline=False,
                ),
            ],
            source_file=Path("/test/email.msg"),
        )

    def test_extract_attachments(self, converter, mock_email, tmp_path):
        """Test attachment extraction."""
        extracted = converter._extract_attachments(mock_email, tmp_path, "test")

        assert len(extracted) == 1
        assert extracted[0].name == "document.pdf"
        assert extracted[0].parent.name == "test_attachments"
        assert extracted[0].read_bytes() == b"PDF_DATA"

    def test_extract_attachments_duplicate_names(self, converter, tmp_path):
        """Test handling of duplicate attachment filenames."""
        email = EmailData(
            subject="Test",
            sender="test@example.com",
            sender_email="test@example.com",
            to=["recipient@example.com"],
            cc=[],
            date=datetime.now(),
            body_html="",
            body_text="",
            attachments=[
                Attachment(
                    filename="file.txt",
                    content_type="text/plain",
                    data=b"First file",
                    is_inline=False,
                ),
                Attachment(
                    filename="file.txt",
                    content_type="text/plain",
                    data=b"Second file",
                    is_inline=False,
                ),
            ],
        )

        extracted = converter._extract_attachments(email, tmp_path, "test")

        assert len(extracted) == 2
        filenames = [f.name for f in extracted]
        assert "file.txt" in filenames
        assert "file_1.txt" in filenames

    def test_extract_attachments_no_attachments(self, converter, mock_email, tmp_path):
        """Test extraction with no file attachments."""
        mock_email.attachments = []
        extracted = converter._extract_attachments(mock_email, tmp_path, "test")
        assert extracted == []

    def test_extract_attachments_inline_only(self, converter, tmp_path):
        """Test that inline attachments are not extracted."""
        email = EmailData(
            subject="Test",
            sender="test@example.com",
            sender_email="test@example.com",
            to=["recipient@example.com"],
            cc=[],
            date=datetime.now(),
            body_html="",
            body_text="",
            attachments=[
                Attachment(
                    filename="image.png",
                    content_type="image/png",
                    data=b"PNG_DATA",
                    content_id="img001",
                    is_inline=True,
                ),
            ],
        )

        extracted = converter._extract_attachments(email, tmp_path, "test")
        assert extracted == []
