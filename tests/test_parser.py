"""Tests for MSG parser."""

import pytest
from pathlib import Path

from msg2pdf.core.parser import MSGParser
from msg2pdf.core.exceptions import MSGParseError


class TestMSGParser:
    """Tests for MSGParser class."""

    def test_parse_nonexistent_file(self):
        """Test parsing a file that doesn't exist."""
        parser = MSGParser()
        with pytest.raises(MSGParseError, match="File not found"):
            parser.parse("/nonexistent/file.msg")

    def test_parse_non_msg_file(self, tmp_path):
        """Test parsing a non-MSG file."""
        text_file = tmp_path / "test.txt"
        text_file.write_text("This is not an MSG file")

        parser = MSGParser()
        with pytest.raises(MSGParseError, match="Not an MSG file"):
            parser.parse(text_file)

    def test_parse_invalid_msg_file(self, tmp_path):
        """Test parsing an invalid MSG file."""
        fake_msg = tmp_path / "fake.msg"
        fake_msg.write_bytes(b"This is not a valid MSG file")

        parser = MSGParser()
        with pytest.raises(MSGParseError, match="Failed to open MSG file"):
            parser.parse(fake_msg)

    def test_parse_recipients_semicolon(self):
        """Test parsing recipients separated by semicolon."""
        parser = MSGParser()
        result = parser._parse_recipients("alice@example.com; bob@example.com")
        assert result == ["alice@example.com", "bob@example.com"]

    def test_parse_recipients_comma(self):
        """Test parsing recipients separated by comma."""
        parser = MSGParser()
        result = parser._parse_recipients("alice@example.com, bob@example.com")
        assert result == ["alice@example.com", "bob@example.com"]

    def test_parse_recipients_empty(self):
        """Test parsing empty recipients."""
        parser = MSGParser()
        assert parser._parse_recipients(None) == []
        assert parser._parse_recipients("") == []

    def test_embed_inline_images(self):
        """Test embedding inline images in HTML."""
        parser = MSGParser()
        html = '<img src="cid:image001">'

        from msg2pdf.core.models import Attachment
        attachments = [
            Attachment(
                filename="image.png",
                content_type="image/png",
                data=b"PNG_DATA",
                content_id="image001",
                is_inline=True,
            )
        ]

        result = parser._embed_inline_images(html, attachments)
        assert "data:image/png;base64," in result
        assert "cid:image001" not in result
