"""PDF rendering using WeasyPrint."""

from pathlib import Path

from jinja2 import Environment, PackageLoader, select_autoescape
from weasyprint import HTML

from msg2pdf.core.exceptions import PDFRenderError
from msg2pdf.core.models import EmailData


class PDFRenderer:
    """Renders email data to PDF using WeasyPrint and Jinja2 templates."""

    def __init__(self):
        """Initialize the renderer with Jinja2 environment."""
        self.env = Environment(
            loader=PackageLoader("msg2pdf", "templates"),
            autoescape=select_autoescape(["html", "xml"]),
        )

    def render(
        self,
        email: EmailData,
        output_path: Path | str,
        show_source: bool = True,
    ) -> Path:
        """Render email data to PDF.

        Args:
            email: Parsed email data.
            output_path: Path for the output PDF file.
            show_source: Whether to show source filename in PDF.

        Returns:
            Path to the generated PDF file.

        Raises:
            PDFRenderError: If rendering fails.
        """
        output_path = Path(output_path)

        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            # Render HTML from template
            html_content = self._render_html(email, show_source)

            # Convert to PDF
            html_doc = HTML(string=html_content)
            html_doc.write_pdf(str(output_path))

            return output_path

        except Exception as e:
            raise PDFRenderError(f"Failed to render PDF: {e}") from e

    def _render_html(self, email: EmailData, show_source: bool) -> str:
        """Render email data to HTML string."""
        template = self.env.get_template("email_full.html")
        return template.render(
            email=email,
            show_source=show_source,
        )

    def render_to_html(
        self,
        email: EmailData,
        output_path: Path | str,
        show_source: bool = True,
    ) -> Path:
        """Render email data to HTML file (for debugging).

        Args:
            email: Parsed email data.
            output_path: Path for the output HTML file.
            show_source: Whether to show source filename.

        Returns:
            Path to the generated HTML file.
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        html_content = self._render_html(email, show_source)
        output_path.write_text(html_content, encoding="utf-8")

        return output_path
