"""CLI commands for msg2pdf."""

from pathlib import Path

import click

from msg2pdf import __version__
from msg2pdf.batch.processor import BatchProcessor
from msg2pdf.core.converter import MSGToPDFConverter
from msg2pdf.core.exceptions import MSG2PDFError


@click.group()
@click.version_option(version=__version__, prog_name="msg2pdf")
def cli():
    """Convert Outlook MSG files to PDF."""
    pass


@cli.command()
@click.argument("msg_file", type=click.Path(exists=True, path_type=Path))
@click.option(
    "-o",
    "--output",
    type=click.Path(path_type=Path),
    default=None,
    help="Output directory (default: same as input file)",
)
@click.option(
    "--no-merge",
    is_flag=True,
    help="Don't merge attachments into PDF (email only)",
)
@click.option(
    "--no-source",
    is_flag=True,
    help="Don't show source filename in PDF",
)
def convert(msg_file: Path, output: Path | None, no_merge: bool, no_source: bool):
    """Convert a single MSG file to PDF.

    Attachments (PDFs, images) are merged into the output PDF by default.

    Example:
        msg2pdf convert email.msg -o ./output/
    """
    # Default output to same directory as input
    if output is None:
        output = msg_file.parent

    converter = MSGToPDFConverter()

    try:
        pdf_path = converter.convert(
            msg_file,
            output,
            merge_attachments=not no_merge,
            show_source=not no_source,
        )
        click.echo(f"✓ Created: {pdf_path}")

    except MSG2PDFError as e:
        click.echo(f"✗ Error: {e}", err=True)
        raise SystemExit(1)


@cli.command()
@click.argument("input_dir", type=click.Path(exists=True, file_okay=False, path_type=Path))
@click.option(
    "-o",
    "--output",
    type=click.Path(path_type=Path),
    required=True,
    help="Output directory for PDFs",
)
@click.option(
    "-r",
    "--recursive",
    is_flag=True,
    help="Search subdirectories for MSG files",
)
@click.option(
    "-w",
    "--workers",
    type=int,
    default=4,
    help="Number of concurrent workers (default: 4)",
)
@click.option(
    "--no-merge",
    is_flag=True,
    help="Don't merge attachments into PDFs (email only)",
)
@click.option(
    "--no-source",
    is_flag=True,
    help="Don't show source filename in PDFs",
)
def batch(
    input_dir: Path,
    output: Path,
    recursive: bool,
    workers: int,
    no_merge: bool,
    no_source: bool,
):
    """Batch convert MSG files in a directory to PDF.

    Attachments (PDFs, images) are merged into each output PDF by default.

    Example:
        msg2pdf batch ./emails/ -o ./pdfs/ --recursive --workers 4
    """
    processor = BatchProcessor(
        max_workers=workers,
        merge_attachments=not no_merge,
        show_source=not no_source,
    )

    # Find MSG files
    msg_files = processor.find_msg_files(input_dir, recursive=recursive)

    if not msg_files:
        click.echo(f"No MSG files found in {input_dir}")
        return

    click.echo(f"Found {len(msg_files)} MSG file(s)")

    # Process with progress bar
    with click.progressbar(
        length=len(msg_files),
        label="Converting",
        show_pos=True,
    ) as progress:

        def on_progress(file: Path, success: bool, error: str | None):
            progress.update(1)

        result = processor.process(msg_files, output, progress_callback=on_progress)

    # Report results
    click.echo()
    click.echo(f"Completed: {result.success_count}/{result.total} successful")

    if result.failed:
        click.echo()
        click.echo("Failed files:")
        for file_path, error in result.failed:
            click.echo(f"  ✗ {file_path.name}: {error}")

        raise SystemExit(1)


@cli.command()
@click.argument("msg_file", type=click.Path(exists=True, path_type=Path))
def info(msg_file: Path):
    """Display information about an MSG file without converting.

    Example:
        msg2pdf info email.msg
    """
    converter = MSGToPDFConverter()

    try:
        email = converter.parse_only(msg_file)

        click.echo(f"Subject:  {email.subject}")
        click.echo(f"From:     {email.sender}")
        if email.sender_email and email.sender_email != email.sender:
            click.echo(f"          <{email.sender_email}>")
        click.echo(f"To:       {email.recipients_display}")
        if email.cc_display:
            click.echo(f"Cc:       {email.cc_display}")
        click.echo(f"Date:     {email.date_display}")
        click.echo()

        # Body info
        if email.body_html:
            click.echo(f"Body:     HTML ({len(email.body_html):,} chars)")
        elif email.body_text:
            click.echo(f"Body:     Plain text ({len(email.body_text):,} chars)")
        else:
            click.echo("Body:     (empty)")

        # Attachments
        if email.attachments:
            click.echo()
            click.echo(f"Attachments ({len(email.attachments)}):")
            for att in email.attachments:
                inline_marker = " [inline]" if att.is_inline else ""
                msg_marker = " [embedded MSG]" if att.is_msg else ""
                click.echo(f"  • {att.filename} ({att.size_display}){inline_marker}{msg_marker}")

    except MSG2PDFError as e:
        click.echo(f"Error: {e}", err=True)
        raise SystemExit(1)


if __name__ == "__main__":
    cli()
