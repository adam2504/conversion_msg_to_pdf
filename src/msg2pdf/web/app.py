"""FastAPI web application for MSG to PDF conversion."""

import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, HTTPException, UploadFile, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from msg2pdf.core.converter import MSGToPDFConverter
from msg2pdf.core.exceptions import MSG2PDFError

app = FastAPI(
    title="MSG to PDF Converter",
    description="Convert Outlook MSG files to PDF",
    version="0.1.0",
)

# Temporary storage for converted files
TEMP_DIR = Path(tempfile.gettempdir()) / "msg2pdf_web"
TEMP_DIR.mkdir(exist_ok=True)

# Store job info (in production, use Redis or database)
jobs: dict[str, dict] = {}


def cleanup_old_files():
    """Remove files older than 1 hour."""
    cutoff = datetime.now() - timedelta(hours=1)
    for job_dir in TEMP_DIR.iterdir():
        if job_dir.is_dir():
            try:
                mtime = datetime.fromtimestamp(job_dir.stat().st_mtime)
                if mtime < cutoff:
                    shutil.rmtree(job_dir)
            except Exception:
                pass


@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main upload page."""
    return """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MSG to PDF Converter</title>
    <style>
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }

        .container {
            background: white;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            max-width: 600px;
            width: 100%;
        }

        h1 {
            color: #333;
            margin-bottom: 8px;
            font-size: 28px;
        }

        .subtitle {
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }

        .upload-area {
            border: 2px dashed #ddd;
            border-radius: 12px;
            padding: 40px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #fafafa;
        }

        .upload-area:hover, .upload-area.dragover {
            border-color: #667eea;
            background: #f0f0ff;
        }

        .upload-area svg {
            width: 48px;
            height: 48px;
            color: #999;
            margin-bottom: 16px;
        }

        .upload-area p {
            color: #666;
            margin-bottom: 8px;
        }

        .upload-area .hint {
            font-size: 12px;
            color: #999;
        }

        input[type="file"] {
            display: none;
        }

        .file-list {
            margin-top: 20px;
            max-height: 200px;
            overflow-y: auto;
        }

        .file-item {
            display: flex;
            align-items: center;
            padding: 10px 12px;
            background: #f5f5f5;
            border-radius: 8px;
            margin-bottom: 8px;
            font-size: 14px;
        }

        .file-item .name {
            flex: 1;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .file-item .size {
            color: #999;
            margin-left: 12px;
        }

        .file-item .remove {
            margin-left: 12px;
            color: #999;
            cursor: pointer;
            font-size: 18px;
        }

        .file-item .remove:hover {
            color: #e74c3c;
        }

        .btn {
            display: inline-block;
            padding: 14px 32px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            margin-top: 20px;
            width: 100%;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }

        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }

        .progress {
            margin-top: 20px;
            display: none;
        }

        .progress-bar {
            height: 8px;
            background: #eee;
            border-radius: 4px;
            overflow: hidden;
        }

        .progress-bar-fill {
            height: 100%;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            width: 0%;
            transition: width 0.3s;
        }

        .progress-text {
            text-align: center;
            margin-top: 8px;
            font-size: 14px;
            color: #666;
        }

        .result {
            margin-top: 20px;
            padding: 16px;
            border-radius: 8px;
            display: none;
        }

        .result.success {
            background: #d4edda;
            color: #155724;
        }

        .result.error {
            background: #f8d7da;
            color: #721c24;
        }

        .download-btn {
            display: inline-block;
            margin-top: 12px;
            padding: 10px 20px;
            background: #28a745;
            color: white;
            text-decoration: none;
            border-radius: 6px;
            font-weight: 600;
        }

        .download-btn:hover {
            background: #218838;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>MSG to PDF Converter</h1>
        <p class="subtitle">Convert Outlook email files (.msg) to PDF with attachments merged</p>

        <div class="upload-area" id="uploadArea">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <p>Drop MSG files here or click to browse</p>
            <p class="hint">You can select multiple files</p>
            <input type="file" id="fileInput" multiple accept=".msg">
        </div>

        <div class="file-list" id="fileList"></div>

        <button class="btn" id="convertBtn" disabled>Convert to PDF</button>

        <div class="progress" id="progress">
            <div class="progress-bar">
                <div class="progress-bar-fill" id="progressFill"></div>
            </div>
            <p class="progress-text" id="progressText">Converting...</p>
        </div>

        <div class="result" id="result"></div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const convertBtn = document.getElementById('convertBtn');
        const progress = document.getElementById('progress');
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        const result = document.getElementById('result');

        let selectedFiles = [];

        // Click to upload
        uploadArea.addEventListener('click', () => fileInput.click());

        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        // File input change
        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
        });

        function handleFiles(files) {
            for (const file of files) {
                if (file.name.toLowerCase().endsWith('.msg')) {
                    if (!selectedFiles.find(f => f.name === file.name)) {
                        selectedFiles.push(file);
                    }
                }
            }
            updateFileList();
        }

        function updateFileList() {
            fileList.innerHTML = '';
            selectedFiles.forEach((file, index) => {
                const div = document.createElement('div');
                div.className = 'file-item';
                div.innerHTML = `
                    <span class="name">${file.name}</span>
                    <span class="size">${formatSize(file.size)}</span>
                    <span class="remove" onclick="removeFile(${index})">&times;</span>
                `;
                fileList.appendChild(div);
            });
            convertBtn.disabled = selectedFiles.length === 0;
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            updateFileList();
        }

        function formatSize(bytes) {
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
            return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
        }

        // Convert button
        convertBtn.addEventListener('click', async () => {
            if (selectedFiles.length === 0) return;

            convertBtn.disabled = true;
            progress.style.display = 'block';
            result.style.display = 'none';
            progressFill.style.width = '0%';
            progressText.textContent = 'Uploading...';

            const formData = new FormData();
            selectedFiles.forEach(file => {
                formData.append('files', file);
            });

            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });

                progressFill.style.width = '50%';
                progressText.textContent = 'Converting...';

                const data = await response.json();

                if (response.ok) {
                    progressFill.style.width = '100%';
                    progressText.textContent = 'Done!';

                    result.className = 'result success';
                    result.style.display = 'block';

                    if (data.file_count === 1) {
                        result.innerHTML = `
                            <strong>Conversion successful!</strong><br>
                            <a href="/download/${data.job_id}" class="download-btn">Download PDF</a>
                        `;
                    } else {
                        result.innerHTML = `
                            <strong>${data.file_count} files converted successfully!</strong><br>
                            <a href="/download/${data.job_id}" class="download-btn">Download ZIP</a>
                        `;
                    }

                    // Clear file list
                    selectedFiles = [];
                    updateFileList();
                } else {
                    throw new Error(data.detail || 'Conversion failed');
                }
            } catch (error) {
                result.className = 'result error';
                result.style.display = 'block';
                result.innerHTML = `<strong>Error:</strong> ${error.message}`;
            }

            convertBtn.disabled = false;
        });
    </script>
</body>
</html>
"""


@app.post("/convert")
async def convert_files(
    background_tasks: BackgroundTasks,
    files: Annotated[list[UploadFile], File(description="MSG files to convert")],
):
    """Convert uploaded MSG files to PDF."""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    # Validate files
    for file in files:
        if not file.filename.lower().endswith(".msg"):
            raise HTTPException(
                status_code=400,
                detail=f"Invalid file type: {file.filename}. Only .msg files are accepted.",
            )

    # Create job directory
    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    input_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)

    # Save uploaded files
    saved_files = []
    for file in files:
        file_path = input_dir / file.filename
        with open(file_path, "wb") as f:
            content = await file.read()
            f.write(content)
        saved_files.append(file_path)

    # Convert files
    converter = MSGToPDFConverter()
    converted = []
    errors = []

    for msg_path in saved_files:
        try:
            pdf_path = converter.convert(
                msg_path,
                output_dir,
                merge_attachments=True,
                show_source=False,
            )
            converted.append(pdf_path)
        except MSG2PDFError as e:
            errors.append(f"{msg_path.name}: {str(e)}")
        except Exception as e:
            errors.append(f"{msg_path.name}: {str(e)}")

    if not converted:
        shutil.rmtree(job_dir)
        raise HTTPException(
            status_code=500,
            detail=f"All conversions failed: {'; '.join(errors)}",
        )

    # Store job info
    jobs[job_id] = {
        "created": datetime.now(),
        "file_count": len(converted),
        "output_dir": output_dir,
        "files": [p.name for p in converted],
    }

    # Schedule cleanup
    background_tasks.add_task(cleanup_old_files)

    return {
        "job_id": job_id,
        "file_count": len(converted),
        "files": [p.name for p in converted],
        "errors": errors if errors else None,
    }


@app.get("/download/{job_id}")
async def download_result(job_id: str):
    """Download converted PDF(s)."""
    if job_id not in jobs:
        # Check if directory exists (for restarts)
        job_dir = TEMP_DIR / job_id / "output"
        if not job_dir.exists():
            raise HTTPException(status_code=404, detail="Job not found or expired")

        # Rebuild job info
        pdf_files = list(job_dir.glob("*.pdf"))
        if not pdf_files:
            raise HTTPException(status_code=404, detail="No files found")

        jobs[job_id] = {
            "created": datetime.now(),
            "file_count": len(pdf_files),
            "output_dir": job_dir,
            "files": [p.name for p in pdf_files],
        }

    job = jobs[job_id]
    output_dir = job["output_dir"]

    if job["file_count"] == 1:
        # Single file - return PDF directly
        pdf_path = output_dir / job["files"][0]
        return FileResponse(
            pdf_path,
            media_type="application/pdf",
            filename=job["files"][0],
        )
    else:
        # Multiple files - create ZIP
        zip_path = output_dir.parent / "converted.zip"
        if not zip_path.exists():
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename in job["files"]:
                    pdf_path = output_dir / filename
                    zf.write(pdf_path, filename)

        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename="converted_emails.zip",
        )


@app.get("/health")
async def health():
    """Health check endpoint."""
    return {"status": "ok"}
