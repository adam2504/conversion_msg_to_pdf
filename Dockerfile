# Use Python with stable Debian (bookworm)
FROM python:3.12-slim-bookworm

# Install WeasyPrint system dependencies
RUN apt-get update && apt-get install -y \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy project files
COPY pyproject.toml .
COPY src/ src/

# Install the package
RUN pip install --no-cache-dir -e .

# Expose port
EXPOSE 8000

# Run the server
CMD ["uvicorn", "msg2pdf.web.app:app", "--host", "0.0.0.0", "--port", "8000"]
