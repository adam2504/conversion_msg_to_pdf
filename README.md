# MSG to PDF Converter

Convert Outlook `.msg` files to PDF, preserving email content, formatting, and merging attachments into a single PDF.

## Features

- Convert single MSG files or batch process entire directories
- **Merge PDF/image attachments** into the output PDF (one file per email)
- Preserve email formatting (HTML body rendered to PDF)
- Embed inline images in the email body
- Concurrent batch processing with configurable workers
- **Web interface** for easy use (no command line needed)
- Cross-platform (macOS, Windows, Linux)

## Installation

### Prerequisites (macOS)

WeasyPrint requires system libraries for PDF rendering:

```bash
brew install pango glib gobject-introspection
```

### Install from source

```bash
pip install -e ".[web]"
```

### Troubleshooting (macOS with Conda)

If you see an error like `cannot load library 'libgobject-2.0-0'`, WeasyPrint cannot find the Homebrew libraries. Add this to your shell profile (`.zshrc` or `.bashrc`):

```bash
export DYLD_LIBRARY_PATH="/opt/homebrew/lib:$DYLD_LIBRARY_PATH"
```

## Web Interface

The easiest way to use the converter - no command line needed:

```bash
./run_server.sh
```

Then open **http://localhost:8000** in your browser.

Features:
- Drag & drop MSG files
- Convert multiple files at once
- Download single PDF or ZIP archive

## CLI Usage

### Convert a single file

```bash
msg2pdf convert email.msg -o output/
```

### Batch convert a directory

```bash
msg2pdf batch ./emails/ -o ./pdfs/ --recursive --workers 4
```

### Inspect an MSG file

```bash
msg2pdf info email.msg
```

## CLI Reference

### `msg2pdf convert`

Convert a single MSG file to PDF with attachments merged.

| Option | Description |
|--------|-------------|
| `-o, --output` | Output directory (default: same as input) |
| `--no-merge` | Don't merge attachments into PDF |
| `--no-source` | Don't show source filename in PDF |

### `msg2pdf batch`

Batch convert MSG files in a directory.

| Option | Description |
|--------|-------------|
| `-o, --output` | Output directory (required) |
| `-r, --recursive` | Search subdirectories |
| `-w, --workers` | Number of concurrent workers (default: 4) |
| `--no-merge` | Don't merge attachments into PDFs |
| `--no-source` | Don't show source filename in PDFs |

### `msg2pdf info`

Display information about an MSG file without converting.

## Attachment Handling

| Type | Action |
|------|--------|
| Inline images (cid:) | Embedded as base64 in email body |
| PDF attachments | Merged after email content |
| Image attachments | Converted to PDF pages and merged |
| Other file types | Listed in email body |

## Development

```bash
# Install with dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Run web server
uvicorn msg2pdf.web.app:app --reload
```

## License

MIT
