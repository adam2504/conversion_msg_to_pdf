"""MSG file parser using extract-msg library."""

import base64
import mimetypes
from datetime import datetime
from pathlib import Path

import extract_msg

from msg2pdf.core.exceptions import MSGParseError
from msg2pdf.core.models import Attachment, EmailData


class MSGParser:
    """Parser for Outlook MSG files."""

    def parse(self, msg_path: Path | str) -> EmailData:
        """Parse an MSG file and return structured email data.

        Args:
            msg_path: Path to the MSG file.

        Returns:
            EmailData object with parsed content.

        Raises:
            MSGParseError: If parsing fails.
        """
        msg_path = Path(msg_path)

        if not msg_path.exists():
            raise MSGParseError(f"File not found: {msg_path}")

        if not msg_path.suffix.lower() == ".msg":
            raise MSGParseError(f"Not an MSG file: {msg_path}")

        try:
            msg = extract_msg.Message(str(msg_path))
        except Exception as e:
            raise MSGParseError(f"Failed to open MSG file: {e}") from e

        try:
            return self._extract_email_data(msg, msg_path)
        finally:
            msg.close()

    def _extract_email_data(
        self, msg: extract_msg.Message, source_path: Path
    ) -> EmailData:
        """Extract email data from parsed MSG object."""
        # Parse sender
        sender = msg.sender or "Unknown"
        sender_email = ""
        if hasattr(msg, "senderEmail") and msg.senderEmail:
            sender_email = msg.senderEmail
        elif hasattr(msg, "sender") and msg.sender and "@" in msg.sender:
            sender_email = msg.sender

        # Parse recipients
        to_list = self._parse_recipients(msg.to)
        cc_list = self._parse_recipients(msg.cc) if msg.cc else []

        # Parse date
        date = None
        if msg.date:
            if isinstance(msg.date, datetime):
                date = msg.date
            elif isinstance(msg.date, str):
                date = self._parse_date_string(msg.date)

        # Parse body
        body_html = msg.htmlBody or ""
        body_text = msg.body or ""

        # Handle body encoding
        if isinstance(body_html, bytes):
            body_html = body_html.decode("utf-8", errors="replace")
        if isinstance(body_text, bytes):
            body_text = body_text.decode("utf-8", errors="replace")

        # Parse attachments
        attachments = self._parse_attachments(msg)

        # Process inline images in HTML body
        if body_html:
            body_html = self._embed_inline_images(body_html, attachments)

        return EmailData(
            subject=msg.subject or "(No Subject)",
            sender=sender,
            sender_email=sender_email,
            to=to_list,
            cc=cc_list,
            date=date,
            body_html=body_html,
            body_text=body_text,
            attachments=attachments,
            source_file=source_path,
        )

    def _parse_recipients(self, recipients_str: str | None) -> list[str]:
        """Parse recipients string into list."""
        if not recipients_str:
            return []

        # Split by semicolon or comma
        if ";" in recipients_str:
            parts = recipients_str.split(";")
        else:
            parts = recipients_str.split(",")

        return [r.strip() for r in parts if r.strip()]

    def _parse_date_string(self, date_str: str) -> datetime | None:
        """Try to parse various date formats."""
        formats = [
            "%a, %d %b %Y %H:%M:%S %z",
            "%d %b %Y %H:%M:%S %z",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S",
            "%m/%d/%Y %I:%M:%S %p",
        ]

        for fmt in formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue

        return None

    def _parse_attachments(self, msg: extract_msg.Message) -> list[Attachment]:
        """Parse attachments from MSG file."""
        attachments = []

        if not hasattr(msg, "attachments") or not msg.attachments:
            return attachments

        for att in msg.attachments:
            try:
                attachment = self._parse_single_attachment(att)
                if attachment:
                    attachments.append(attachment)
            except Exception:
                # Skip problematic attachments
                continue

        return attachments

    def _parse_single_attachment(self, att) -> Attachment | None:
        """Parse a single attachment."""
        # Get filename
        filename = None
        if hasattr(att, "longFilename") and att.longFilename:
            filename = att.longFilename
        elif hasattr(att, "shortFilename") and att.shortFilename:
            filename = att.shortFilename
        elif hasattr(att, "filename") and att.filename:
            filename = att.filename

        if not filename:
            return None

        # Get data
        data = None
        if hasattr(att, "data") and att.data:
            data = att.data
        elif hasattr(att, "save"):
            # Some attachments need to be read differently
            try:
                data = att.data
            except Exception:
                return None

        if not data:
            return None

        # Determine content type
        content_type = "application/octet-stream"
        if hasattr(att, "mimetype") and att.mimetype:
            content_type = att.mimetype
        else:
            guessed = mimetypes.guess_type(filename)[0]
            if guessed:
                content_type = guessed

        # Check if it's an inline image
        content_id = None
        is_inline = False
        if hasattr(att, "contentId") and att.contentId:
            content_id = att.contentId.strip("<>")
            is_inline = True
        elif hasattr(att, "cid") and att.cid:
            content_id = att.cid.strip("<>")
            is_inline = True

        # Check if embedded MSG
        is_msg = filename.lower().endswith(".msg")

        return Attachment(
            filename=filename,
            content_type=content_type,
            data=data if isinstance(data, bytes) else data.encode(),
            content_id=content_id,
            is_inline=is_inline,
            is_msg=is_msg,
        )

    def _embed_inline_images(
        self, html: str, attachments: list[Attachment]
    ) -> str:
        """Replace cid: references with base64 data URIs."""
        for att in attachments:
            if att.is_inline and att.content_id:
                # Create base64 data URI
                b64_data = base64.b64encode(att.data).decode("ascii")
                data_uri = f"data:{att.content_type};base64,{b64_data}"

                # Replace cid: references
                html = html.replace(f"cid:{att.content_id}", data_uri)
                html = html.replace(f"CID:{att.content_id}", data_uri)

        return html
