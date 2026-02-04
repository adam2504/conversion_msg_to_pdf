"""Tests for batch processor."""

import pytest
from pathlib import Path

from msg2pdf.batch.processor import BatchProcessor, BatchResult


class TestBatchResult:
    """Tests for BatchResult class."""

    def test_empty_result(self):
        """Test empty batch result."""
        result = BatchResult()
        assert result.total == 0
        assert result.success_count == 0
        assert result.failure_count == 0
        assert result.success_rate == 0.0

    def test_all_successful(self):
        """Test all successful conversions."""
        result = BatchResult(
            successful=[Path("/output/a.pdf"), Path("/output/b.pdf")],
            failed=[],
        )
        assert result.total == 2
        assert result.success_count == 2
        assert result.failure_count == 0
        assert result.success_rate == 100.0

    def test_all_failed(self):
        """Test all failed conversions."""
        result = BatchResult(
            successful=[],
            failed=[
                (Path("/input/a.msg"), "Error 1"),
                (Path("/input/b.msg"), "Error 2"),
            ],
        )
        assert result.total == 2
        assert result.success_count == 0
        assert result.failure_count == 2
        assert result.success_rate == 0.0

    def test_mixed_results(self):
        """Test mixed success and failure."""
        result = BatchResult(
            successful=[Path("/output/a.pdf")],
            failed=[(Path("/input/b.msg"), "Error")],
        )
        assert result.total == 2
        assert result.success_count == 1
        assert result.failure_count == 1
        assert result.success_rate == 50.0


class TestBatchProcessor:
    """Tests for BatchProcessor class."""

    def test_find_msg_files_empty_dir(self, tmp_path):
        """Test finding MSG files in empty directory."""
        processor = BatchProcessor()
        files = processor.find_msg_files(tmp_path)
        assert files == []

    def test_find_msg_files(self, tmp_path):
        """Test finding MSG files."""
        # Create test files
        (tmp_path / "a.msg").touch()
        (tmp_path / "b.msg").touch()
        (tmp_path / "c.txt").touch()

        processor = BatchProcessor()
        files = processor.find_msg_files(tmp_path)

        assert len(files) == 2
        assert all(f.suffix == ".msg" for f in files)

    def test_find_msg_files_recursive(self, tmp_path):
        """Test recursive MSG file finding."""
        # Create nested structure
        subdir = tmp_path / "subdir"
        subdir.mkdir()

        (tmp_path / "a.msg").touch()
        (subdir / "b.msg").touch()

        processor = BatchProcessor()

        # Non-recursive
        files = processor.find_msg_files(tmp_path, recursive=False)
        assert len(files) == 1

        # Recursive
        files = processor.find_msg_files(tmp_path, recursive=True)
        assert len(files) == 2

    def test_find_msg_files_invalid_dir(self):
        """Test finding MSG files in non-existent directory."""
        processor = BatchProcessor()
        with pytest.raises(ValueError, match="Not a directory"):
            processor.find_msg_files("/nonexistent/dir")
