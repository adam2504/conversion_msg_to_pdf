#!/bin/bash
# Start the MSG to PDF web server

# Set library path for WeasyPrint on macOS
export DYLD_LIBRARY_PATH="/opt/homebrew/lib:$DYLD_LIBRARY_PATH"

echo "Starting MSG to PDF Converter..."
echo "Open http://localhost:8000 in your browser"
echo ""

uvicorn msg2pdf.web.app:app --host 0.0.0.0 --port 8000
