"""Batch processing for MSG to PDF conversion."""

from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from msg2pdf.core.converter import MSGToPDFConverter


@dataclass
class BatchResult:
    """Result of batch processing."""

    successful: list[Path] = field(default_factory=list)
    failed: list[tuple[Path, str]] = field(default_factory=list)

    @property
    def total(self) -> int:
        """Total files processed."""
        return len(self.successful) + len(self.failed)

    @property
    def success_count(self) -> int:
        """Number of successful conversions."""
        return len(self.successful)

    @property
    def failure_count(self) -> int:
        """Number of failed conversions."""
        return len(self.failed)

    @property
    def success_rate(self) -> float:
        """Success rate as percentage."""
        if self.total == 0:
            return 0.0
        return (self.success_count / self.total) * 100


class BatchProcessor:
    """Process multiple MSG files concurrently."""

    def __init__(
        self,
        max_workers: int = 4,
        merge_attachments: bool = True,
        show_source: bool = True,
    ):
        """Initialize batch processor.

        Args:
            max_workers: Maximum number of concurrent workers.
            merge_attachments: Whether to merge attachments into PDF.
            show_source: Whether to show source in PDFs.
        """
        self.max_workers = max_workers
        self.merge_attachments = merge_attachments
        self.show_source = show_source
        self.converter = MSGToPDFConverter()

    def find_msg_files(
        self,
        input_dir: Path | str,
        recursive: bool = False,
    ) -> list[Path]:
        """Find all MSG files in a directory.

        Args:
            input_dir: Directory to search.
            recursive: Whether to search subdirectories.

        Returns:
            List of MSG file paths.
        """
        input_dir = Path(input_dir)

        if not input_dir.is_dir():
            raise ValueError(f"Not a directory: {input_dir}")

        pattern = "**/*.msg" if recursive else "*.msg"
        return sorted(input_dir.glob(pattern))

    def process(
        self,
        msg_files: list[Path],
        output_dir: Path | str,
        progress_callback: Callable[[Path, bool, str | None], None] | None = None,
    ) -> BatchResult:
        """Process multiple MSG files.

        Args:
            msg_files: List of MSG file paths.
            output_dir: Output directory for PDFs.
            progress_callback: Optional callback for progress updates.
                Called with (file_path, success, error_message).

        Returns:
            BatchResult with successful and failed conversions.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        result = BatchResult()

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            future_to_file = {
                executor.submit(
                    self._process_single,
                    msg_file,
                    output_dir,
                ): msg_file
                for msg_file in msg_files
            }

            # Collect results as they complete
            for future in as_completed(future_to_file):
                msg_file = future_to_file[future]

                try:
                    pdf_path = future.result()
                    result.successful.append(pdf_path)

                    if progress_callback:
                        progress_callback(msg_file, True, None)

                except Exception as e:
                    error_msg = str(e)
                    result.failed.append((msg_file, error_msg))

                    if progress_callback:
                        progress_callback(msg_file, False, error_msg)

        return result

    def _process_single(self, msg_file: Path, output_dir: Path) -> Path:
        """Process a single MSG file.

        Args:
            msg_file: Path to MSG file.
            output_dir: Output directory.

        Returns:
            Path to generated PDF.
        """
        return self.converter.convert(
            msg_file,
            output_dir,
            merge_attachments=self.merge_attachments,
            show_source=self.show_source,
        )
