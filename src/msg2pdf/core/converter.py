"""Main converter orchestrating MSG parsing and PDF rendering."""

import io
import tempfile
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter

from msg2pdf.core.exceptions import AttachmentError, MSG2PDFError
from msg2pdf.core.models import Attachment, EmailData
from msg2pdf.core.parser import MSGParser
from msg2pdf.core.renderer import PDFRenderer


class MSGToPDFConverter:
    """Orchestrates conversion from MSG to PDF."""

    def __init__(self):
        """Initialize converter with parser and renderer."""
        self.parser = MSGParser()
        self.renderer = PDFRenderer()

    def convert(
        self,
        msg_path: Path | str,
        output_dir: Path | str,
        output_filename: str | None = None,
        merge_attachments: bool = True,
        show_source: bool = True,
    ) -> Path:
        """Convert an MSG file to PDF with attachments merged.

        Args:
            msg_path: Path to the MSG file.
            output_dir: Directory for output files.
            output_filename: Optional custom filename for PDF (without extension).
            merge_attachments: Whether to merge attachments into PDF.
            show_source: Whether to show source filename in PDF.

        Returns:
            Path to the generated PDF file.

        Raises:
            MSG2PDFError: If conversion fails.
        """
        msg_path = Path(msg_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # Parse the MSG file
        email = self.parser.parse(msg_path)

        # Determine output filename
        if output_filename:
            pdf_name = f"{output_filename}.pdf"
        else:
            pdf_name = f"{msg_path.stem}.pdf"

        pdf_path = output_dir / pdf_name

        # Render email to PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_email_pdf = Path(tmp.name)

        self.renderer.render(email, tmp_email_pdf, show_source=show_source)

        # Merge attachments if requested
        if merge_attachments and email.file_attachments:
            self._merge_attachments(tmp_email_pdf, email.file_attachments, pdf_path)
            tmp_email_pdf.unlink()  # Clean up temp file
        else:
            # Just move the email PDF to output
            tmp_email_pdf.rename(pdf_path)

        return pdf_path

    def parse_only(self, msg_path: Path | str) -> EmailData:
        """Parse MSG file without converting to PDF.

        Useful for inspection or custom processing.

        Args:
            msg_path: Path to the MSG file.

        Returns:
            Parsed email data.
        """
        return self.parser.parse(msg_path)

    def _merge_attachments(
        self,
        email_pdf_path: Path,
        attachments: list[Attachment],
        output_path: Path,
    ) -> None:
        """Merge attachments into the email PDF.

        Args:
            email_pdf_path: Path to the email PDF.
            attachments: List of attachments to merge.
            output_path: Path for the merged output PDF.
        """
        writer = PdfWriter()

        # Add email PDF pages
        email_reader = PdfReader(str(email_pdf_path))
        for page in email_reader.pages:
            writer.add_page(page)

        # Process each attachment
        for att in attachments:
            try:
                self._add_attachment_to_pdf(writer, att)
            except Exception:
                # Skip attachments that can't be converted
                # They're already listed in the email body
                continue

        # Write merged PDF
        with open(output_path, "wb") as f:
            writer.write(f)

    def _add_attachment_to_pdf(self, writer: PdfWriter, att: Attachment) -> None:
        """Add an attachment to the PDF writer.

        Args:
            writer: PDF writer to add pages to.
            att: Attachment to convert and add.
        """
        filename_lower = att.filename.lower()

        if filename_lower.endswith(".pdf"):
            # Merge PDF attachment
            self._merge_pdf_attachment(writer, att)
        elif any(filename_lower.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"]):
            # Convert image to PDF page
            self._add_image_attachment(writer, att)
        # Other file types are skipped (already listed in email body)

    def _merge_pdf_attachment(self, writer: PdfWriter, att: Attachment) -> None:
        """Merge a PDF attachment into the writer.

        Args:
            writer: PDF writer to add pages to.
            att: PDF attachment.
        """
        try:
            pdf_reader = PdfReader(io.BytesIO(att.data))
            for page in pdf_reader.pages:
                writer.add_page(page)
        except Exception as e:
            raise AttachmentError(f"Failed to merge PDF '{att.filename}': {e}") from e

    def _add_image_attachment(self, writer: PdfWriter, att: Attachment) -> None:
        """Convert an image attachment to a PDF page and add it.

        Args:
            writer: PDF writer to add page to.
            att: Image attachment.
        """
        try:
            # Open image
            img = Image.open(io.BytesIO(att.data))

            # Convert to RGB if necessary (for PNG with transparency, etc.)
            if img.mode in ("RGBA", "LA", "P"):
                # Create white background
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                background.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                img = background
            elif img.mode != "RGB":
                img = img.convert("RGB")

            # Convert to PDF
            pdf_bytes = io.BytesIO()
            img.save(pdf_bytes, format="PDF", resolution=100.0)
            pdf_bytes.seek(0)

            # Add to writer
            img_reader = PdfReader(pdf_bytes)
            for page in img_reader.pages:
                writer.add_page(page)

        except Exception as e:
            raise AttachmentError(f"Failed to convert image '{att.filename}': {e}") from e
